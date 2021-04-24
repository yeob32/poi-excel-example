package com.example.demo.excel

import com.example.demo.excel.resource.ExcelCellKey
import com.example.demo.excel.resource.ExcelRenderResource
import com.example.demo.excel.resource.ExcelRenderResource.Companion.prepareExcelResource
import com.example.demo.excel.resource.converter.ExcelConverter
import com.example.demo.excel.resource.dataformat.DefaultDataFormatDecider
import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import java.io.IOException
import java.io.OutputStream
import java.time.Instant
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*

abstract class AbstractExcelFile<T> {

    companion object {
        protected val supplyExcelVersion = SpreadsheetVersion.EXCEL2007
    }

    private val workbook: SXSSFWorkbook
    private val sheet: Sheet
    private val excelRenderResource: ExcelRenderResource

    protected constructor(type: Class<T>) : this(emptyList(), type)

    protected constructor(data: List<T>, type: Class<T>) {
        validate(data)
        this.workbook = SXSSFWorkbook()
        this.sheet = workbook.createSheet()
        this.excelRenderResource = prepareExcelResource(workbook, type, DefaultDataFormatDecider())
        renderExcel(data)
    }

    protected abstract fun renderExcel(data: List<T>)

    private fun validate(data: List<T>) {
        val maxRows = supplyExcelVersion.maxRows
        require(data.size <= maxRows) {
            "This concrete ExcelFile does not support over $maxRows rows"
        }
    }

    protected fun renderHeaders(rowIndex: Int, cellStartIndex: Int) {
        var cellIndex = cellStartIndex
        sheet.createRow(rowIndex).run {
            excelRenderResource.fields.forEach { fieldName ->
                createCell(cellIndex++).run {
                    setCellValue(excelRenderResource.headerMap[fieldName])
                    cellStyle = excelRenderResource.cellStyleMap.getCellStyleMap(ExcelCellKey.ofHeader(fieldName))
                }
            }
        }
    }

    protected fun renderData(dataList: List<Any>, rowIndex: Int, cellStartIndex: Int) {
        var index = rowIndex
        dataList.forEach { data ->
            var cellIndex = cellStartIndex
            sheet.createRow(index++).run {
                excelRenderResource.fields.forEach { fieldName ->
                    val field = data.javaClass.declaredFields
                        .firstOrNull { it.name == fieldName } ?: throw RuntimeException("Invalid data")
                    field.isAccessible = true
                    createCell(cellIndex++).run {
                        cellStyle = excelRenderResource.cellStyleMap.getCellStyleMap(ExcelCellKey.ofBody(fieldName))
                        setCellValue(this, field.get(data), excelRenderResource.converterMap[fieldName]!!)
                    }
                }
            }
        }
    }

    private fun setCellValue(cell: Cell, value: Any?, converter: ExcelConverter) = when (value) {
        is Number -> cell.setCellValue(value.toDouble())
        is Instant -> cell.setCellValue(Date.from(value))
        is Date -> cell.setCellValue(value)
        is LocalDate -> cell.setCellValue(value)
        is LocalDateTime -> cell.setCellValue(value)
        else -> cell.setCellValue(converter.convert(value?.toString() ?: ""))
    }

    @Throws(IOException::class)
    fun write(stream: OutputStream) {
        workbook.write(stream)
        workbook.close()
        workbook.dispose()
        stream.close()
    }

}
