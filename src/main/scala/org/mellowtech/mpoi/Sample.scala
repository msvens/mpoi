package org.mellowtech.mpoi

/**
 * Created by msvens on 03/09/15.
 */

import java.io.FileOutputStream

import org.apache.poi.ss.usermodel.{Workbook => PWorkbook, Sheet => PSheet, Row => PRow, Cell => PCell}

object Sample extends App {

  import XSSFConverter._
  val f = Font(bold = Some(true))
  val s = CellStyle(font = Some(f))

  val l = List(SCell(0, "some value"), BCell(1, true))


  val r = Row(0, l, Some(s))
  val wb = Workbook(
    List(
      Sheet(Some("sheet1"), List(r), Some(Map(0 -> s))),
      Sheet(Some("sheet2"), Nil)
    )
  )

  val pwb: PWorkbook = wb

  val fileOut = new FileOutputStream("/Users/msvens/temp/workbook.xlsx");
  pwb.write(fileOut);
  fileOut.close();

}
