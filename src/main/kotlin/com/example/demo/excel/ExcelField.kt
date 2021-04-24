package com.example.demo.excel

import com.example.demo.excel.resource.converter.DefaultExcelConverter
import kotlin.reflect.KClass

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelField(
    val headerName: String,
    val converter: KClass<*> = DefaultExcelConverter::class
)