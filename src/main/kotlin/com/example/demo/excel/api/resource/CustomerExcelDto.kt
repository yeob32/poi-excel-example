package com.example.demo.excel.api.resource

import com.example.demo.excel.ExcelBodyStyle
import com.example.demo.excel.ExcelField
import com.example.demo.excel.ExcelHeaderStyle
import com.example.demo.excel.ExcelResource
import org.apache.poi.ss.usermodel.IndexedColors
import java.time.Instant

@ExcelBodyStyle
@ExcelHeaderStyle(fontBold = true, fontColor = IndexedColors.BLUE)
@ExcelResource("회원 목록")
data class CustomerExcelDto(

    @ExcelField("Name") val name: String,
    @ExcelField("Address") val address: String,
    @ExcelField val age: Int,

    val test: Long,

    @ExcelField("날짜") val createdAt: Instant,

    val price1: Float,
    val price2: Double,
    val price3: Long,
)
