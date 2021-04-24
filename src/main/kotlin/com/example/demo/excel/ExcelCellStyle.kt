package com.example.demo.excel

import org.apache.poi.ss.usermodel.IndexedColors

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelCellStyle(
  val fontBold: Boolean = false,
  val fontColor: IndexedColors = IndexedColors.BLACK1,
)
