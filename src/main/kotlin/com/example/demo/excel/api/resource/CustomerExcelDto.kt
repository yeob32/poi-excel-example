package com.example.demo.excel.api.resource

import com.example.demo.customer.Status
import com.example.demo.excel.*
import com.example.demo.excel.resource.converter.DefaultExcelConverter
import com.example.demo.excel.resource.converter.ExcelConverter
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

    @ExcelField("상태", converter = StatusConverter::class) val status: Status,

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

class StatusConverter : ExcelConverter {
    override fun convert(value: String) = when (Status.valueOf(value)) {
        Status.ACTIVE -> "활동중"
        Status.INACTIVE -> "활동중지"
    }
}