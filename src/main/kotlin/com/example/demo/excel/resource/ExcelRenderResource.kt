package com.example.demo.excel.resource

import com.example.demo.excel.resource.dataformat.DefaultDataFormatDecider
import com.example.demo.excel.ExcelBodyStyle
import com.example.demo.excel.ExcelField
import com.example.demo.excel.ExcelHeaderStyle
import com.example.demo.excel.ExcelResource
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Workbook
import java.lang.RuntimeException

class ExcelRenderResource(
    val fields: List<String>,
    val headerMap: MutableMap<String, String>,
    val cellStyleMap: ExcelCellStyleMap
) {
    companion object {
        fun prepareExcelResource(
            workbook: Workbook,
            type: Class<*>,
            dataFormatDecider: DefaultDataFormatDecider
        ): ExcelRenderResource {
            if (!type.isAnnotationPresent(ExcelResource::class.java)) {
                throw RuntimeException("is not ExcelResource : ${type.name}")
            }

            val excelField = ExcelField::class.java
            val excelHeaderStyle = ExcelHeaderStyle::class.java
            val excelBodyStyle = ExcelBodyStyle::class.java

            val headerMap = mutableMapOf<String, String>()
            val fields = mutableListOf<String>()

            val excelCellStyleMap = ExcelCellStyleMap(dataFormatDecider)

            val hasDefaultExcelHeaderStyle = type.isAnnotationPresent(excelHeaderStyle)
            val hasDefaultExcelBodyStyle = type.isAnnotationPresent(excelBodyStyle)

            type.declaredFields
                .forEach { field ->
                    val fieldName = field.name

                    fields.add(fieldName)
                    headerMap[fieldName] = when {
                        field.isAnnotationPresent(excelField) -> field.getAnnotation(excelField)
                            .headerName.ifEmpty { fieldName }
                        else -> fieldName
                    }

                    if (hasDefaultExcelHeaderStyle) {
                        excelCellStyleMap.put(
                            workbook,
                            field.type,
                            ExcelCellKey.of(fieldName, ExcelRenderLocation.HEADER),
                            getHeaderStyles(workbook, type.getAnnotation(excelHeaderStyle))
                        )
                    }

                    if (hasDefaultExcelBodyStyle) {
                        excelCellStyleMap.put(
                            workbook,
                            field.type,
                            ExcelCellKey.of(fieldName, ExcelRenderLocation.BODY),
                            getBodyStyles(workbook, type.getAnnotation(excelBodyStyle))
                        )
                    }
                }

            return ExcelRenderResource(fields, headerMap, excelCellStyleMap)
        }

        private fun getBodyStyles(workbook: Workbook, annotation: ExcelBodyStyle): CellStyle {
            val headerFont = workbook.createFont()
            val headerCellStyle = workbook.createCellStyle()
            headerCellStyle.setFont(headerFont)
            return headerCellStyle
        }

        private fun getHeaderStyles(workbook: Workbook, annotation: ExcelHeaderStyle): CellStyle {
            val headerFont = workbook.createFont()
            headerFont.bold = annotation.fontBold
            headerFont.color = annotation.fontColor.getIndex()
            val headerCellStyle = workbook.createCellStyle()
            headerCellStyle.setFont(headerFont)
            return headerCellStyle
        }
    }
}