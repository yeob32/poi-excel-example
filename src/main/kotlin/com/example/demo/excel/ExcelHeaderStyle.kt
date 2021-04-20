package com.example.demo.excel

import org.apache.poi.ss.usermodel.IndexedColors

@Target(AnnotationTarget.CLASS, AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelHeaderStyle(
    val fontBold: Boolean = true,
    val fontColor: IndexedColors = IndexedColors.BLACK1
)
