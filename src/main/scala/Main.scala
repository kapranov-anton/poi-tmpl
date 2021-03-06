package kaa.poi

import kaa.poi.xlsx.Replace
import kaa.poi.model._

import java.io.FileOutputStream
object Main extends App {
  val in: ReplacementDict = Map(
    "companyFrom.name" -> Text("COMPANY NAME"),
    "company.name" -> Text("COMPANY NAME"),
    "process.createdAt" -> Table(List(
      Row(List(Cell("1"), Cell(""), Cell("xxx"), Cell("YYY"))),
      Row(List(Cell("2"), Cell(""), Cell("XXX"), Cell("yyy"))))),
    "billItems.name" -> Table(List(
      Row(List(Cell("1"), Cell(""), Cell("xxx"), Cell("YYY"))),
      Row(List(Cell("2"), Cell(""), Cell("XXX"), Cell("yyy")))))
  ).withDefaultValue(Text("-< ??? >-"))
  //val result = Replace("application_rosalco_legal.docx", in)
  val result = Replace("bill_template.xlsx", in)
  val os = new FileOutputStream("/tmp/result.xlsx")
  os.write(result)
}
