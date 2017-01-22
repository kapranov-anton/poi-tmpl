organization := "kaa"

name := "poi-tmpl"

version := "0.1-SNAPSHOT"

scalaVersion := "2.11.8"

val poiV = "3.15"
libraryDependencies ++= Seq(
  "org.apache.poi" % "poi"       % poiV
, "org.apache.poi" % "poi-ooxml" % poiV
)
