package com.example.demo.excel.resource.dataformat

import org.apache.poi.ss.usermodel.DataFormat

interface DataFormatDecider {

    fun getDataFormat(dataFormat: DataFormat, type: Class<*>): Short
}