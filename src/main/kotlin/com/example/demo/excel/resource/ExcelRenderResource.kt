package com.example.demo.excel.resource

import com.example.demo.excel.resource.dataformat.DefaultDataFormatDecider
import com.example.demo.excel.ExcelBodyStyle
import com.example.demo.excel.ExcelField
import com.example.demo.excel.ExcelHeaderStyle
import com.example.demo.excel.ExcelResource
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Workbook

class ExcelRenderResource(
    val fields: List<String>,
    val headerMap: MutableMap<String, String>,
    val cellStyleMap: ExcelCellStyleMap
) {
    companion object {
        fun prepareExcelResource(workbook: Workbook, type: Class<*>): ExcelRenderResource {
            val excelResource = ExcelResource::class.java
            val excelField = ExcelField::class.java
            val excelHeaderStyle = ExcelHeaderStyle::class.java
            val excelBodyStyle = ExcelBodyStyle::class.java

            val excelCellStyleMap = ExcelCellStyleMap(DefaultDataFormatDecider())

            val headerMap = mutableMapOf<String, String>()
            val fields = mutableListOf<String>()

            val isExcelResource = type.isAnnotationPresent(excelResource)
            val hasExcelHeaderStyle = type.isAnnotationPresent(excelHeaderStyle)
            val hasExcelBodyStyle = type.isAnnotationPresent(excelBodyStyle)

            type.declaredFields
                .filter { field -> isExcelResource || field.isAnnotationPresent(excelField) }
                .forEach { field ->
                    val fieldName = field.name
                    if (field.isAnnotationPresent(excelField)) {
                        val headerName = field.getAnnotation(excelField)
                            .headerName.ifEmpty { fieldName }
                        headerMap[fieldName] = headerName
                        fields.add(fieldName)
                    } else {
                        headerMap[fieldName] = fieldName
                        fields.add(fieldName)
                    }

                    if (hasExcelHeaderStyle) {
                        excelCellStyleMap.put(
                            workbook,
                            field.type,
                            ExcelCellKey.of(fieldName, ExcelRenderLocation.HEADER),
                            getHeaderStyles(workbook, type.getAnnotation(excelHeaderStyle))
                        )
                    }

                    if (hasExcelBodyStyle) {
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