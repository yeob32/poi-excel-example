package com.example.demo.excel

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelField(
    val headerName: String
)
