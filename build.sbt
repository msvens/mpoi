name := "mpoi"

version := "0.1-SNAPSHOT"

scalaVersion := "2.11.7"


// Change this to another test framework if you prefer
libraryDependencies += "org.scalatest" %% "scalatest" % "2.2.5" % "test"
libraryDependencies += "com.typesafe.scala-logging" %% "scala-logging" % "3.1.0"

//Apache POI
libraryDependencies += "org.apache.poi" % "poi-ooxml" % "3.12"