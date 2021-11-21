import sbt.Keys.libraryDependencies

name := "scala-poi"

version := "0.1"

scalaVersion := "2.12.8"

libraryDependencies ++= Seq(
  "org.apache.spark" %% "spark-sql" % "2.4.4",
  "org.apache.commons" % "commons-text" % "1.9",
  "org.scalatest" %% "scalatest" % "3.2.2" % "test",
  "org.apache.poi" % "poi-ooxml-full" % "5.0.0",
  "org.apache.poi" % "poi-ooxml" % "5.0.0"
)
