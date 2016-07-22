package org.mellowtech.mpoi

import java.util.{Calendar, Date}

/**
 * Created by msvens on 03/09/15.
 */
sealed trait Cell{
  def idx: Int
  def style: Option[CellStyle]
}

case class FCell(idx: Int, data: String, style: Option[CellStyle] = None) extends Cell
case class SCell(idx: Int, data: String, style: Option[CellStyle] = None) extends Cell
case class NCell(idx: Int, data: Double, style: Option[CellStyle] = None) extends Cell
case class BCell(idx: Int, data: Boolean, style: Option[CellStyle] = None) extends Cell
case class DCell(idx: Int, data: Date, style: Option[CellStyle] = None) extends Cell
case class CCell(idx: Int, data: Calendar, style: Option[CellStyle] = None) extends Cell
case class ECell(idx: Int, style: Option[CellStyle] = None) extends Cell{
  def data = Unit
}

object Cell {
  def toCell[A](a: A, idx: Int): Cell = a match {
    case x: String => x match {
      case x if x == null => SCell(idx, x)
      case x if x.startsWith("=") => FCell(idx, x.substring(1))
      case _ => SCell(idx, x)
    }
    case x: Boolean => BCell(idx, x)
    case x: Date => println("this is a date"); DCell(idx, x)
    case x: Calendar => CCell(idx, x)
    case x: java.lang.Number => NCell(idx, x.doubleValue())
    case _ => SCell(idx, a.toString())

  }
}

case class Row(idx: Int, cells: Seq[Cell], style: Option[CellStyle] = None)

case class Sheet(name: Option[String] = None, rows: Seq[Row], styles: Option[Map[Int, CellStyle]] = None)

case class Workbook(sheets: Seq[Sheet]){
  import org.apache.poi.ss.usermodel.{Workbook => PWorkbook, Sheet => PSheet, Row => PRow, Cell => PCell, CellStyle => PCellStyle}

  var styles: Map[CellStyle, PCellStyle] = Map()

  def getStyle(style: CellStyle) : PCellStyle = styles.getOrElse(style, {
    null
  })
}

object Workbook{

  def simplified(header: Boolean, data: Seq[Seq[Any]]): Workbook = {
    val headerStyle = header match {
      case true => Some(CellStyle(font = Some(Font(bold = Some(true)))))
      case false => None
    }

    val first = List(Row(0, data.head.zipWithIndex.map{case (data,idx) => Cell.toCell(data,idx)}, headerStyle))


    val rs = data.tail.foldLeft(first){(l,row) =>
      val rr = Row(l.head.idx + 1,row.zipWithIndex.map{case (data,idx) => Cell.toCell(data, idx)},None)
      rr :: l
    }

    Workbook(List(Sheet(rows = rs)))




    /*for(row <- data; cell <- row){
      cell match {
        case x: String => println("string")
        case x: java.lang.Number => println("number")
        case _ => println("unknown type")
      }
    }
    null*/
  }
}

object Align extends Enumeration {
  type Align = Value
  val CENTER, CENTER_SELECTION, FILL, GENERAL, JUSTIFY, LEFT, RIGHT = Value
}

/**
 * Wrapper for POI font. Most properties map directly to the corresponding Font properties. All properties are
 * optional
 * @see [[https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/Font.html Poi Font]]
 * @see [[https://poi.apache.org/apidocs/org/apache/poi/hssf/record/IndexRecord.html Indexed Color]]
 * @param name font name
 * @param bold true for bold
 * @param italic true for italic
 * @param color use color codes from Apache POI (see Poi Font above)
 * @param underline use undrline codes from Poi Font
 *
 */
case class Font(name: Option[String] = None,
                bold: Option[Boolean] = None,
                italic: Option[Boolean] = None,
                color: Option[Short] = None,
underline: Option[Byte] = None)


import Align._
case class CellStyle(align: Option[Align] = None, font: Option[Font])

