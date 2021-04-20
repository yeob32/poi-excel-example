package com.example.demo.excel.resource

import com.example.demo.excel.resource.dataformat.DataFormatDecider
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.Workbook

class ExcelCellStyleMap(private val dataFormatDecider: DataFormatDecider) {

    private val cellStyleMap = mutableMapOf<ExcelCellKey, CellStyle>()

    fun put(workbook: Workbook, type: Class<*>, cellKey: ExcelCellKey, cellStyle: CellStyle) {
        val dataFormat: DataFormat = workbook.createDataFormat()
        cellStyle.dataFormat = dataFormatDecider.getDataFormat(dataFormat, type)
        cellStyleMap[cellKey] = cellStyle
    }

    fun getCellStyleMap(key: ExcelCellKey): CellStyle? {
        return cellStyleMap[key]
    }
}