package com.example.demo.excel.api

import com.example.demo.excel.api.resource.CustomerExcelDto
import com.example.demo.customer.Customer
import com.example.demo.customer.CustomerRepository
import com.example.demo.excel.api.resource.CustomerDto
import com.example.demo.excel.reader.ExcelReader
import com.example.demo.excel.resource.ExcelReaderResource
import com.example.demo.excel.sheet.SimpleExcelFile
import org.springframework.web.bind.annotation.*
import org.springframework.web.multipart.MultipartFile
import java.io.File
import java.io.FileInputStream
import javax.servlet.http.HttpServletResponse

@RequestMapping(value = ["/api/excel"])
@RestController
class ExcelApi(private val customerRepository: CustomerRepository) {

    @GetMapping("/customers.xlsx")
    fun customerReport(httpServletResponse: HttpServletResponse) {
        httpServletResponse.contentType = "ms-vnd/excel"
        httpServletResponse.setHeader("Content-Disposition", "attachment;filename=test.xls")

        val customers = customerRepository.findAll().toDto()
        val excel = SimpleExcelFile(customers, CustomerExcelDto::class.java)
        excel.write(httpServletResponse.outputStream)
    }

    @PostMapping("/read")
    fun readExcel(@RequestParam("file") multipartFile: MultipartFile) {
        val a = ExcelReaderResource.prepareExcelResource(CustomerDto::class.java)
        val clazz = CustomerDto::class.java
        val cons = clazz.getConstructor(*a.types.toTypedArray())

        cons.newInstance(*arrayListOf("1234", 1234).toTypedArray())

        val file =
            FileInputStream(File("/Users/ksy/IdeaProjects/kotlin-spring-excel/src/main/resources/static/customers.xlsx"))
        val test = ExcelReader(file, CustomerDto::class.java).read()

        println(test)
    }

    fun List<Customer>.toDto() = map { it.toDto() }
    fun Customer.toDto() = CustomerExcelDto(
        name = name,
        address = address,
        age = age,
        test = 0,
        createdAt = createdAt,
        price1 = 1234567.01230f,
        price2 = 1234567.01230,
        price3 = 1234567,
        status = status
    )
}