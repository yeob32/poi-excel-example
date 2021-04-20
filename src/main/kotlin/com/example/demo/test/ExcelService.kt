package com.example.demo.test

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.NumberToTextConverter
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile

@Service
class ExcelService {

    fun readExcel(multipartFile: MultipartFile): SpreadSheet {
        WorkbookFactory.create(multipartFile.inputStream).use {
            val sheet = it.getSheetAt(0)

            val rows = arrayListOf<List<String>>()
            val header = extractRow(sheet.getRow(0))

            for (rowIdx in 1..sheet.lastRowNum) {
                rows.add(extractRow(sheet.getRow(rowIdx)))
            }

            return SpreadSheet(multipartFile.name, header, rows)
        }
    }

    fun readExcelTest(multipartFile: MultipartFile): SpreadSheetMap {
        WorkbookFactory.create(multipartFile.inputStream).use {
            val sheet = it.getSheetAt(0)

            val rows = arrayListOf<Map<String, String>>()
            val header = extractHeader(sheet.getRow(0))

            for (rowIdx in 1..sheet.lastRowNum) {
                rows.add(extractRow(header, sheet.getRow(rowIdx)))
            }

            return SpreadSheetMap(multipartFile.originalFilename!!, header, rows)
        }
    }

    private fun extractRow(row: Row): List<String> {
        val rows = arrayListOf<String>()
        for (idx in row.firstCellNum until row.lastCellNum) {
            val cell = row.getCell(idx)
            if (row.getCell(idx) == null) {
                continue
            }

            val cellValue = when (cell.cellType) {
                CellType.NUMERIC -> NumberToTextConverter.toText(cell.numericCellValue)
                else -> {
                    cell.stringCellValue
                }
            }

            rows.add(cellValue)
        }

        return rows
    }

    private fun extractHeader(row: Row): List<String> {
        val rows = arrayListOf<String>()
        for (idx in row.firstCellNum until row.lastCellNum) {
            val cell = row.getCell(idx)
            if (row.getCell(idx) == null) {
                continue
            }

            val cellValue = when (cell.cellType) {
                CellType.NUMERIC -> NumberToTextConverter.toText(cell.numericCellValue)
                else -> {
                    cell.stringCellValue
                }
            }

            rows.add(cellValue)
        }

        return rows
    }

    private fun extractRow(header: List<String>, row: Row): HashMap<String, String> {
        val rowMap = hashMapOf<String, String>()
        for (idx in row.firstCellNum until row.lastCellNum) {
            val cell = row.getCell(idx)
            if (row.getCell(idx) == null) {
                continue
            }

            val cellValue = when (cell.cellType) {
                CellType.NUMERIC -> NumberToTextConverter.toText(cell.numericCellValue)
                else -> {
                    cell.stringCellValue
                }
            }

            rowMap[header[idx]] = cellValue
        }

        return rowMap
    }
}