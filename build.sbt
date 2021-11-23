/**
 * Copyright 2021 Synkre Consulting, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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
