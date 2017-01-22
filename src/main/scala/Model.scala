package kaa.poi.model

import scala.util.matching.Regex.Match


case class Cell(value: String)
case class Row(value: List[Cell])

abstract class Replacement
case class Text(value: String) extends Replacement
case class Table(value: List[Row]) extends Replacement

trait Replacer {
  val tmplVariablePattern = "\\$\\{[^\\}]+\\}".r
  def unwrap(s: String) = s.stripPrefix("${").stripSuffix("}")
  def replaceText(src: String, dict: ReplacementDict): String =
    tmplVariablePattern.replaceAllIn(src, _ match { case Match(s) =>
      dict(unwrap(s)) match {
        case Text(s) => s
        case _ => "" // never reached
      }
    })

  def findTable(src: String, dict: ReplacementDict): Option[List[Row]] =
    tmplVariablePattern.findAllIn(src).flatMap { m =>
      dict(unwrap(m)) match {
        case Table(t) => List(t)
        case _ => List()
      }
    }.toList.headOption
}
