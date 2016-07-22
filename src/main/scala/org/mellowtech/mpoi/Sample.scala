package org.mellowtech.mpoi

/**
 * Created by msvens on 03/09/15.
 */

import java.io.FileOutputStream
import java.util.Date

import org.apache.poi.ss.usermodel.{Workbook => PWorkbook, Sheet => PSheet, Row => PRow, Font => PFont, Cell => PCell, IndexedColors}

object Sample extends App {

  import XSSFConverter._
  val f = Font(bold = Some(true), color = Option(IndexedColors.BLUE.index), underline = Some(PFont.U_DOUBLE))
  val s = CellStyle(font = Some(f))

  val l = List(SCell(0, "some value"), BCell(1, true))


  val r = Row(0, l)
  val wb = Workbook(
    List(
      Sheet(Some("sheet1"), List(r), Some(Map(0 -> s))),
      Sheet(Some("sheet2"), Nil)
    )
  )

  val pwb: PWorkbook = wb

  val fileOut = new FileOutputStream("/Users/msvens/temp/workbook.xlsx")
  pwb.write(fileOut)
  fileOut.close()

  //Try simplified
  //val headers = List("number", "string", "formula")
  val data = Seq(
      List("number", "string", "formula", "date"),
      List(1, "string", "=1+1", new Date())
  )

  val fileOut1 = new FileOutputStream("/Users/msvens/temp/workbook1.xlsx")
  Workbook.simplified(true,data).write(fileOut1)
  fileOut.close()


}
