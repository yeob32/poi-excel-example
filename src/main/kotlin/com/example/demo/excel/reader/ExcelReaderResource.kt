package com.example.demo.excel.reader

import com.example.demo.excel.ExcelField
import com.example.demo.excel.ExcelResource
import java.lang.RuntimeException

data class PropertyInfo(
    val name: String,
    val type: Class<*>,
)

class ExcelReaderResource(
    val fields: List<String>,
    val types: List<Class<*>>,
    val propertyMap: MutableMap<String, PropertyInfo>
) {
    companion object {
        fun prepareExcelResource(
            type: Class<*>,
        ): ExcelReaderResource {
            check(type.isAnnotationPresent(ExcelResource::class.java)) {
                throw RuntimeException("is not ExcelResource : ${type.name}")
            }

            val excelField = ExcelField::class.java
            val headerMap = mutableMapOf<String, String>()
            val types = mutableListOf<Class<*>>()
            val fields = arrayListOf<String>()
            val propertyMap = mutableMapOf<String, PropertyInfo>()

            type.declaredFields
                .filter { field -> field.isAnnotationPresent(excelField) }
                .forEach { field ->
                    val headerName = field.getAnnotation(excelField).headerName
                    fields.add(headerName)
                    types.add(field.type)
                    headerMap[headerName] = field.name
                    propertyMap[headerName] = PropertyInfo(field.name, field.type)
                }

            return ExcelReaderResource(fields, types, propertyMap)
        }
    }
}