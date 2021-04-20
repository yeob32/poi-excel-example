package com.example.demo.excel.sheet

import com.example.demo.excel.AbstractExcelFile

@Suppress("UNCHECKED_CAST")
class SimpleExcelFile<T>(data: List<T>, type: Class<T>): AbstractExcelFile<T>(data, type) {

    companion object {
        private const val ROW_START_INDEX = 0
        private const val COLUMN_START_INDEX = 0
        private var currentRowIndex = ROW_START_INDEX
    }

    override fun renderExcel(data: List<T>) {
        renderHeaders(currentRowIndex++, COLUMN_START_INDEX)

        if(data.isNotEmpty()) {
            renderData(data as List<Any>, currentRowIndex++, COLUMN_START_INDEX)
        }
    }
}