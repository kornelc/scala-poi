package com.synkre.poi

import java.io.FileOutputStream

import org.apache.poi.xssf.usermodel.{XSSFTable, XSSFWorkbook}
import org.apache.spark.sql.functions.col
import org.apache.spark.sql.{DataFrame, SparkSession}

import scala.collection.JavaConverters._
import scala.util.Try

object Excel {

  import java.io.FileInputStream

  private def readWorkbook(path: String): Try[XSSFWorkbook] = Try{
    new XSSFWorkbook(new FileInputStream(path))
  }

  private def copyData(workbook: XSSFWorkbook, dfs: Map[String, DataFrame]): Try[Unit] = Try{
    for( name <- workbook.getAllNames.asScala){
      println(s"$name formula ${name.getRefersToFormula}")
    }
    for((tablename, df) <- dfs;
        table <- Option(workbook.getTable(tablename))
          .toRight(new RuntimeException(s"Table not found: $tablename"))
          .toTry) {
      println(s"Replacing $tablename")
      val dfWithOrderedFields = df.select(table.getColumns.asScala.map(n => col(n.getName)): _*)
      dfWithOrderedFields.show(1000)
      val data = dfWithOrderedFields.collect().map(_.toSeq)
      copyData(table, data)
    }

  }
  private def copyData(table: XSSFTable, data: Array[Seq[_]]): Try[Unit] = Try {
    val sheet = table.getXSSFSheet

    val oldTotalRowIndex = table.getEndRowIndex
    table.setDataRowCount(data.length)

    val tableWidth = data(table.getStartRowIndex + 1).length -1
    val templateCellFormats =
      (for(c <- 0 to tableWidth)
        yield sheet.getRow(0).getCell(c)).toList
    templateCellFormats.foreach(f => println(s"FORMATS: $f"))

    val totalCellFormulas =
      (for(c <- 0 to tableWidth)
        yield Try(Option(sheet.getRow(oldTotalRowIndex).getCell(c).getCellFormula)).getOrElse(None)).toList
    totalCellFormulas.foreach(f => println(s"FORMULAS: $f"))

    for (r <- 0 to data.length-1) {
      val row = sheet.createRow(table.getStartRowIndex + r +1);
      for(c <- 0 to data(0).length -1) {
        val cell = row.createCell(c)
        cell.setCellType(templateCellFormats(c).getCellType)
        cell.setCellStyle(templateCellFormats(c).getCellStyle)
        val field = data(r)(c)
        if (field.isInstanceOf[String]) {
          cell.setCellValue(field.asInstanceOf[String])
        } else if (field.isInstanceOf[Double]) {
          cell.setCellValue(field.asInstanceOf[Double])
        }
      }
    }

    val totalRowIndex = table.getStartRowIndex + data.length + 1
    val totalRow = sheet.createRow(totalRowIndex)

    for {
      formulaOpt <- totalCellFormulas.zipWithIndex
      formula <- formulaOpt._1
      column = formulaOpt._2
      _ = println(s"Creating total cell in row $totalRowIndex col $column")
      totalCell = totalRow.createCell(column)
      _ = totalCell.setCellFormula(formula)
      _ = println(s"Written total row in $totalRowIndex $column $formula")
    }
    ()
  }

  def replaceRegions(dfs: Map[String, DataFrame], templatePath: String, targetSpreadSheetPath: String)(implicit spark: SparkSession): Either[Throwable, XSSFWorkbook] = {

    val result = for {
      workbook <- readWorkbook(templatePath)
      _ <- copyData(workbook, dfs)
      _ = workbook.write(new FileOutputStream(targetSpreadSheetPath))
    } yield workbook
    result.failed.foreach(_.printStackTrace())
    result.toEither
  }
}
