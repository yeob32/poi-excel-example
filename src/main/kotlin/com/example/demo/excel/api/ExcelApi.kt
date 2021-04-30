package com.example.demo.excel.api

import com.example.demo.excel.api.resource.CustomerExcelDto
import com.example.demo.customer.Customer
import com.example.demo.customer.CustomerRepository
import com.example.demo.customer.Status
import com.example.demo.excel.api.resource.CustomerDto
import com.example.demo.excel.reader.ExcelReader
import com.example.demo.excel.reader.ExcelReaderResource
import com.example.demo.excel.sheet.SimpleExcelFile
import org.springframework.web.bind.annotation.*
import org.springframework.web.multipart.MultipartFile
import java.io.File
import java.io.FileInputStream
import java.time.Instant
import javax.servlet.http.HttpServletResponse

@RequestMapping(value = ["/api/excel"])
@RestController
class ExcelApi(private val customerRepository: CustomerRepository) {

    @GetMapping("/customers")
    fun customer(httpServletResponse: HttpServletResponse) {
        httpServletResponse.contentType = "ms-vnd/excel"
        httpServletResponse.setHeader("Content-Disposition", "attachment;filename=test.xls")

        val customers = arrayListOf<CustomerExcelDto>()
        for (i in 0..500000L) {
            customers.add(
                CustomerExcelDto(
                    name = "test_$i",
                    address = "test_$i",
                    age = i.toInt(),
                    test = i,
                    status = Status.INACTIVE,
                    createdAt = Instant.now(),
                    price2 = i.toDouble(),
                    price1 = i.toFloat(),
                    price3 = i
                )
            )
        }

        val excel = SimpleExcelFile(customers, CustomerExcelDto::class.java)
        excel.write(httpServletResponse.outputStream)
    }

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

    private fun List<Customer>.toDto() = map { CustomerExcelDto.of(it) }
}