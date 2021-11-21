package com.synkre.poi

import java.nio.file.{Files, Paths}

import org.apache.spark.sql.SparkSession
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class ExcelSpec extends AnyFlatSpec with Matchers {
  implicit val spark = SparkSession
    .builder()
    .master("local[*]")
    .appName("scala-poi-test")
    .getOrCreate()

  "Excel" should "copy data into excel sheet from dataframes" in {
    val outputPath = "/tmp/out.xlsm"
    Files.deleteIfExists(Paths.get(outputPath))
    val newWorkbookEither = Excel.replaceRegions(
      Map(
        "SALES" -> spark
                  .read
                  .option("ignoreTrailingWhiteSpace", true)
                  .json("src/test/resources/excel/sales.json")
      ),
      "src/test/resources/excel/SALES.xlsm",
      outputPath)
    newWorkbookEither.isRight shouldEqual true
    newWorkbookEither.foreach{ wb =>
      wb.getTable("SALES").getEndRowIndex shouldEqual 7
    }
  }
}
