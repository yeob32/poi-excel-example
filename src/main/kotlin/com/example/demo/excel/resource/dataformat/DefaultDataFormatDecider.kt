package com.example.demo.excel.resource.dataformat

import org.apache.poi.ss.usermodel.DataFormat
import java.time.Instant
import java.time.LocalDateTime

class DefaultDataFormatDecider : DataFormatDecider {

    companion object {
        private const val CURRENT_FORMAT = "#,##0"
        private const val FLOAT_FORMAT_2_DECIMAL_PLACES = "#,##0.00"
        private const val DEFAULT_FORMAT = ""
        private const val DEFAULT_DATE_FORMAT = "m/d/yy h:mm"
    }

    override fun getDataFormat(dataFormat: DataFormat, type: Class<*>): Short = when {
        isFloatType(type) -> dataFormat.getFormat(FLOAT_FORMAT_2_DECIMAL_PLACES)
        isIntegerType(type) -> dataFormat.getFormat(CURRENT_FORMAT)
        isDateType(type) -> dataFormat.getFormat(DEFAULT_DATE_FORMAT)
        else -> dataFormat.getFormat(DEFAULT_FORMAT)
    }

    private fun isIntegerType(type: Class<*>): Boolean = listOf<Class<*>?>(
        Byte::class.java, Byte::class.javaPrimitiveType,
        Short::class.java, Short::class.javaPrimitiveType,
        Int::class.java, Int::class.javaPrimitiveType,
        Long::class.java, Long::class.javaPrimitiveType
    ).contains(type)

    private fun isFloatType(type: Class<*>): Boolean {
        val floatTypes = listOf<Class<*>?>(
            Float::class.java, Float::class.javaPrimitiveType,
            Double::class.java, Double::class.javaPrimitiveType
        )
        return floatTypes.contains(type)
    }

    private fun isBooleanType(type: Class<*>): Boolean = listOf<Class<*>?>(
        Boolean::class.java, Boolean::class.javaPrimitiveType
    ).contains(type)

    private fun isDateType(type: Class<*>) = listOf<Class<*>?>(
        Instant::class.java, LocalDateTime::class.java
    ).contains(type)
}