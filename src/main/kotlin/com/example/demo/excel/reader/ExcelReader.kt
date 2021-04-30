package com.example.demo.excel.reader

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.NumberToTextConverter
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream

class ExcelReader<T> {

    private val sheet: Sheet
    private val excelReaderResource: ExcelReaderResource
    private val type: Class<T>

    constructor(inputStream: InputStream, type: Class<T>) {
        this.sheet = XSSFWorkbook(inputStream).getSheetAt(0)
        this.excelReaderResource = ExcelReaderResource.prepareExcelResource(type)
        this.type = type
    }

    fun read(): List<T> {
        val headers = readHeaders()
        val sheetRows = arrayListOf<Map<String, String>>()
        for (rowIdx in 1..sheet.lastRowNum) {
            sheetRows.add(extractRow(headers, sheet.getRow(rowIdx)))
        }

        val cons = type.getConstructor(*excelReaderResource.types.toTypedArray())
        return sheetRows.map { rowMap ->
            val values = excelReaderResource.fields.map { fieldKey ->
                rowMap[fieldKey]?.run {
                    getValue(this, excelReaderResource.propertyMap[fieldKey]!!.type)
                }
            }

            cons.newInstance(*values.toTypedArray())
        }
    }

    fun getValue(value: String, type: Class<*>) = when (type) {
        String::class.java -> value
        Int::class.java -> value.toInt()
        Long::class.java -> value.toLong()
        Float::class.java -> value.toFloat()
        Double::class.java -> value.toDouble()
        Boolean::class.java -> value.toBoolean()
        else -> value
    }

    private fun readHeaders() = extractRow(sheet.getRow(0))

    private fun extractRow(row: Row): List<String> {
        val rows = arrayListOf<String>()
        for (idx in row.firstCellNum until row.lastCellNum) {
            row.getCell(idx)?.run {
                val cellValue = when (cellType) {
                    CellType.NUMERIC -> NumberToTextConverter.toText(numericCellValue)
                    else -> stringCellValue
                }

                rows.add(cellValue)
            }
        }

        return rows
    }

    private fun extractRow(header: List<String>, row: Row): HashMap<String, String> {
        val rowMap = hashMapOf<String, String>()
        for (idx in row.firstCellNum until row.lastCellNum) {
            row.getCell(idx)?.run {
                val cellValue = when (cellType) {
                    CellType.NUMERIC -> NumberToTextConverter.toText(numericCellValue)
                    else -> stringCellValue
                }

                rowMap[header[idx]] = cellValue
            }
        }

        return rowMap
    }
}
