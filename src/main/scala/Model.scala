package kaa.poi.model

case class Cell(value: String)
case class Row(value: List[Cell])

abstract class Replacement
case class Text(value: String) extends Replacement
case class Table(value: List[Row]) extends Replacement

