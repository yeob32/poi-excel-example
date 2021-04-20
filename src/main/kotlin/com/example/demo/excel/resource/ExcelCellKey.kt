package com.example.demo.excel.resource

import org.apache.poi.ss.usermodel.CellStyle

class ExcelCellKey(
    private val dataFieldName: String,
    private val excelRenderLocation: ExcelRenderLocation
) {
    companion object {
        fun of(fieldName: String, excelRenderLocation: ExcelRenderLocation) = ExcelCellKey(fieldName, excelRenderLocation)
        fun ofHeader(fieldName: String) = ExcelCellKey(fieldName, ExcelRenderLocation.HEADER)
        fun ofBody(fieldName: String) = ExcelCellKey(fieldName, ExcelRenderLocation.BODY)
    }

    override fun equals(other: Any?): Boolean {
        if (this === other) return true
        if (javaClass != other?.javaClass) return false

        other as ExcelCellKey

        if (dataFieldName != other.dataFieldName) return false
        if (excelRenderLocation != other.excelRenderLocation) return false

        return true
    }

    override fun hashCode(): Int {
        var result = dataFieldName.hashCode()
        result = 31 * result + excelRenderLocation.hashCode()
        return result
    }

    override fun toString(): String {
        return "ExcelCellKey(dataFieldName='$dataFieldName', excelRenderLocation=$excelRenderLocation)"
    }
}

fun main() {
    println(ExcelCellKey.ofHeader("zzz") == ExcelCellKey.ofBody("zzz"))

    val cellStyleMap = mutableMapOf<ExcelCellKey, String>()
    cellStyleMap[ExcelCellKey.ofHeader("zzz")] = "hihihi"
    cellStyleMap[ExcelCellKey.ofBody("zzz")] = "byebye"

    println(cellStyleMap.size)
}