package com.example.demo.excel.resource

import com.example.demo.excel.*
import com.example.demo.excel.api.resource.CustomerExcelDto
import com.example.demo.excel.resource.dataformat.DefaultDataFormatDecider
import com.example.demo.excel.resource.converter.ExcelConverter
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Workbook
import java.lang.RuntimeException
import kotlin.reflect.KClass

class ExcelRenderResource(
    val fields: List<String>,
    val headerMap: MutableMap<String, String>,
    val converterMap: MutableMap<String, ExcelConverter>,
    val cellStyleMap: ExcelCellStyleMap,
) {
    companion object {
        fun prepareExcelResource(
            workbook: Workbook,
            type: Class<*>,
            dataFormatDecider: DefaultDataFormatDecider,
        ): ExcelRenderResource {
            if (!type.isAnnotationPresent(ExcelResource::class.java)) {
                throw RuntimeException("is not ExcelResource : ${type.name}")
            }

            val excelResource = type.getAnnotation(ExcelResource::class.java)
            val headerMap = mutableMapOf<String, String>()
            val fields = mutableListOf<String>()
            val converterMap = mutableMapOf<String, ExcelConverter>()

            val excelCellStyleMap = ExcelCellStyleMap(dataFormatDecider)

            type.declaredFields
                .filter { field -> field.name != "Companion" }
                .forEach { field ->
                    check(field.isAnnotationPresent(ExcelField::class.java)) {
                        throw RuntimeException("does not ExcelField : ${field.name}")
                    }

                    val fieldName = field.name
                    fields.add(fieldName)

                    if (field.isAnnotationPresent(ExcelField::class.java)) {
                        val excelField = field.getAnnotation(ExcelField::class.java)
                        headerMap[fieldName] = excelField.headerName
                        converterMap[fieldName] = createObject(excelField.converter) as ExcelConverter
                    }

                    excelCellStyleMap.put(
                        workbook,
                        field.type,
                        ExcelCellKey.ofHeader(fieldName),
                        getHeaderStyles(workbook, excelResource.defaultHeaderStyle)
                    )

                    excelCellStyleMap.put(
                        workbook,
                        field.type,
                        ExcelCellKey.ofBody(fieldName),
                        getBodyStyles(workbook, excelResource.defaultBodyStyle)
                    )
                }

            return ExcelRenderResource(fields, headerMap, converterMap, excelCellStyleMap)
        }

        private fun getBodyStyles(workbook: Workbook, annotation: ExcelCellStyle): CellStyle {
            val headerFont = workbook.createFont()
            val headerCellStyle = workbook.createCellStyle()
            headerCellStyle.setFont(headerFont)
            return headerCellStyle
        }

        private fun getHeaderStyles(workbook: Workbook, annotation: ExcelCellStyle): CellStyle {
            val headerFont = workbook.createFont()
            headerFont.bold = annotation.fontBold
            headerFont.color = annotation.fontColor.getIndex()
            val headerCellStyle = workbook.createCellStyle()
            headerCellStyle.setFont(headerFont)
            return headerCellStyle
        }
    }
}

inline fun <reified T : Any> createObject(clazz: KClass<out T>, vararg args: Any?): T {
    return clazz.constructors.first().call(*args)
}