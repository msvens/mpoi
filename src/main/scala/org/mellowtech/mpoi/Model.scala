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

case class Row(idx: Int, cells: Seq[Cell], style: Option[CellStyle] = None)

case class Sheet(name: Option[String], rows: Seq[Row], styles: Option[Map[Int, CellStyle]] = None)

case class Workbook(sheets: Seq[Sheet]){
  import org.apache.poi.ss.usermodel.{Workbook => PWorkbook, Sheet => PSheet, Row => PRow, Cell => PCell, CellStyle => PCellStyle}

  var styles: Map[CellStyle, PCellStyle] = Map()

  def getStyle(style: CellStyle) : PCellStyle = styles.getOrElse(style, {
    null
  })
}

object Align extends Enumeration {
  type Align = Value
  val CENTER, CENTER_SELECTION, FILL, GENERAL, JUSTIFY, LEFT, RIGHT = Value
}

case class Font(name: Option[String] = None, bold: Option[Boolean] = None, italic: Option[Boolean] = None)


import Align._
case class CellStyle(align: Option[Align] = None, font: Option[Font])

