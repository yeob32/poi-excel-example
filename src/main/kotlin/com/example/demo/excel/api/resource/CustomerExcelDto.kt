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
    @ExcelField("Age") val age: Int,

    @ExcelField("Test") val test: Long,

    @ExcelField("날짜") val createdAt: Instant,

    @ExcelField("Price1") val price1: Float,
    @ExcelField("Price2") val price2: Double,
    @ExcelField("Price3") val price3: Long
)

@ExcelResource("asdf")
data class CustomerDto(
    @ExcelField("Name")  val name: String,
    @ExcelField("Age")  val age: Int
)