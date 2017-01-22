package kaa.poi.xlsx

import scala.util.matching.Regex.Match
import java.io.{FileInputStream, ByteArrayOutputStream}
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.xssf.usermodel.{XSSFWorkbook, XSSFSheet, XSSFRow, XSSFCell}
import org.apache.poi.ss.usermodel.CellCopyPolicy
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
import org.apache.poi.ss.util.CellRangeAddress

object Replace extends Replacer {
  def copyRowStyles(srcRow: XSSFRow, dstRow: XSSFRow) {
    dstRow.setRowStyle(srcRow.getRowStyle)
    dstRow.setHeight(srcRow.getHeight)
    srcRow.cellIterator.toList.foreach { srcCell =>
      val dstCell = dstRow.createCell(srcCell.getColumnIndex)
      dstCell.setCellStyle(srcCell.getCellStyle)
    }
  }

  def getRegions(sheet: XSSFSheet) =
    (0 to sheet.getNumMergedRegions - 1)
      .map { i => (i, sheet.getMergedRegion(i)) }

  def shiftRegion(region: CellRangeAddress, shift: Int) =
    new CellRangeAddress(
      region.getFirstRow + shift,
      region.getLastRow + shift,
      region.getFirstColumn,
      region.getLastColumn)

  def copyRowRegions(srcRow: XSSFRow, dstRow: XSSFRow) {
    val sheet = srcRow.getSheet
    getRegions(sheet)
      .filter { case (_, region) =>
        region.getFirstRow == srcRow.getRowNum && region.getLastRow == srcRow.getRowNum }
      .foreach { case (_, region) =>
        sheet.addMergedRegion(shiftRegion(region, dstRow.getRowNum - srcRow.getRowNum)) }
  }

  def detachRegionsWithShift(sheet: XSSFSheet, shiftFrom: Int, shiftSize: Int): Seq[CellRangeAddress] =
    getRegions(sheet).reverse.map { case (index, region) =>
      sheet.removeMergedRegion(index)
      shiftRegion(region, if (region.getFirstRow > shiftFrom) shiftSize else 0)
    }

  def attachRegions(sheet: XSSFSheet, regions: Seq[CellRangeAddress]) {
    regions.foreach(sheet.addMergedRegion(_))
  }

  def copyRows(cell: XSSFCell, tableRows: List[Row]) {
    val rowCount = tableRows.length
    val sheet = cell.getSheet
    val cellIndex = cell.getColumnIndex
    val sampleRow = cell.getRow
    val sampleRowIndex = cell.getRowIndex
    val firstNewRowIndex = sampleRowIndex + 1
    val lastNewRowIndex = sampleRowIndex + rowCount
    if (rowCount > 0) {
      val savedRegions = detachRegionsWithShift(sheet, sampleRowIndex, rowCount)
      sheet.shiftRows(firstNewRowIndex, sheet.getLastRowNum, rowCount, true, false)
      attachRegions(sheet, savedRegions)

      firstNewRowIndex to lastNewRowIndex foreach { rIndex =>
        val newRow = sheet.createRow(rIndex)
        copyRowStyles(sampleRow, newRow)
        copyRowRegions(sampleRow, newRow)
      }

      for {
        (Row(tRow), tRowIndex) <- tableRows.zipWithIndex
        (Cell(tCell), tCellIndex) <- tRow.zipWithIndex
      } sheet
          .getRow(tRowIndex + sampleRowIndex)
          .getCell(tCellIndex + cellIndex)
          .setCellValue(tCell)
    }
  }

  def apply(fileName: String, dict: ReplacementDict) = {
    val pkg = OPCPackage.open(new FileInputStream(fileName))
    val wb = new XSSFWorkbook(pkg)
    for {
      sheet <- wb.sheetIterator.toList
      row <- sheet.rowIterator.toList
      cell <- row.cellIterator.toList.asInstanceOf[List[XSSFCell]]
    } findTable(cell.toString, dict) match {
      case Some(table) => copyRows(cell, table)
      case None => cell.setCellValue(replaceText(cell.toString, dict))
    }

    val os = new ByteArrayOutputStream()
    wb.write(os)
    pkg.close()
    os.toByteArray()
  }
}
