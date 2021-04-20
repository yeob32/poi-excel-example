package com.example.demo.excel

@Target(AnnotationTarget.CLASS)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelResource(
    val value: String,
)
