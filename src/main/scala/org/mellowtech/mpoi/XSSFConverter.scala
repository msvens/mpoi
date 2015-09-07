package org.mellowtech.mpoi

/**
 * Created by msvens on 03/09/15.
 */

import org.apache.poi.ss.usermodel.{Workbook => PWorkbook, Sheet => PSheet, Row => PRow, Cell => PCell, CellStyle => PCellStyle,Font => PFont}
import org.apache.poi.xssf.usermodel.XSSFWorkbook

object XSSFConverter {

  case class Styler(wb: PWorkbook) {
    var styles: Map[CellStyle, PCellStyle] = Map()

    def style(s: CellStyle): PCellStyle = styles.getOrElse(s, {
      val ps = wb.createCellStyle()
      styles = styles + ((s, ps))
      if (s.font.isDefined) {
        val f = s.font.get
        val pf = wb.createFont()
        if (f.bold.isDefined) pf.setBold(f.bold.get)
        if (f.italic.isDefined) pf.setItalic(f.italic.get)
        if (f.name.isDefined) pf.setFontName(f.name.get)
        ps.setFont(pf)
      }
      ps
    })
  }


  private def createRow(wb: Styler, sheet: PSheet, row: Row): Unit = {
    val r = sheet.createRow(row.idx)
    if(row.style.isDefined) {
      println("setting row style")
      r.setRowStyle(wb.style(row.style.get))
    }
    row.cells foreach {c =>
      val pcell = r.createCell(c.idx)
      if(c.style.isDefined) pcell.setCellStyle(wb.style(c.style.get))
      else if(row.style.isDefined) pcell.setCellStyle(wb.style(row.style.get))
      c match {
        case x: BCell => pcell.setCellValue(x.data)
        case x: SCell => pcell.setCellValue(x.data)
        case x: NCell => pcell.setCellValue(x.data)
        case x: FCell => pcell.setCellFormula(x.data)
        case x: DCell => pcell.setCellValue(x.data)
        case x: CCell => pcell.setCellValue(x.data)
        case x: ECell => pcell.setCellType(PCell.CELL_TYPE_ERROR)
      }
    }
  }

  private def createSheet(wb: Styler, sheet: Sheet): Unit = {
    val s = sheet.name match {
      case None => wb.wb.createSheet()
      case Some(n) => wb.wb.createSheet(n)
    }
    if(sheet.styles.isDefined)
      sheet.styles.get.foreach{case (key, value)  =>
        val ps = wb.style(value)
        println("setting sheet style "+key)
          s.setDefaultColumnStyle(key,ps)
      }

    sheet.rows foreach {r =>
      createRow(wb, s, r)
    }

  }

  implicit def export(wb: Workbook): PWorkbook = {
    val pwb = new XSSFWorkbook()
    val styler = Styler(pwb)
    wb.sheets foreach {s =>
      createSheet(styler, s)
    }
    pwb
  }

}
