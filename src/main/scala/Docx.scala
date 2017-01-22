package kaa.poi.docx

import scala.util.matching.Regex.Match
import java.io.{FileInputStream, ByteArrayOutputStream}
import org.apache.poi.xwpf.usermodel.{XWPFDocument, XWPFParagraph, XWPFTable, XWPFRun}
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth
import scala.collection.JavaConversions._ // deprecated in scala 2.12
import kaa.poi.model.{
  ReplacementDict
, Replacement
, Text
, Table
, Row
, Cell
, Replacer
}

object Replace extends Replacer {
  def removeAllRuns(para: XWPFParagraph) {
    para.getRuns.size to 0 by -1 foreach { para.removeRun(_) }
  }

  def fullWidth(table: XWPFTable) {
    val tblW = table.getCTTbl().getTblPr().getTblW()
    tblW.setW(BigInt(10000).bigInteger)
    tblW.setType(STTblWidth.PCT)
  }

  def insertEmptyTable(para: XWPFParagraph): XWPFTable = {
    val cursor = para.getCTP.newCursor
    val table = para.getBody.insertNewTbl(cursor)
    table.removeRow(0)
    table
  }

  def apply(fileName: String, dict: ReplacementDict) = {
    val doc = new XWPFDocument(new FileInputStream(fileName))

    val tableParas = for {
      table <- doc.getTables.toList
      row <- table.getRows.toList
      cell <- row.getTableCells.toList
      para <- cell.getParagraphs.toList
    } yield para
    
    tableParas ++ doc.getParagraphs.toList foreach { para =>
      findTable(para.getText, dict) match {
        case Some(rows) =>
          removeAllRuns(para)
          val table = insertEmptyTable(para)
          fullWidth(table)

          for {
            Row(row) <- rows
            tableRow = table.createRow
            Cell(cell) <- row
            tableCell = tableRow.createCell
          } tableCell.setText(cell)

        case None =>
          val repl = replaceText(para.getText, dict)
          removeAllRuns(para)
          para.insertNewRun(0).setText(repl)
      }
    }

    val os = new ByteArrayOutputStream()
    doc.write(os)
    os.toByteArray()
  }
}
