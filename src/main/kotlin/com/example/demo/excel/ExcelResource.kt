package com.example.demo.excel

@Target(AnnotationTarget.CLASS)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelResource(
    val value: String,
    val defaultHeaderStyle: ExcelCellStyle = ExcelCellStyle(fontBold = true),
    val defaultBodyStyle: ExcelCellStyle = ExcelCellStyle()
)
