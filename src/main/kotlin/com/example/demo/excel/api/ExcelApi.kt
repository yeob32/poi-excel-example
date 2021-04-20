package com.example.demo.excel.api

import com.example.demo.excel.api.resource.CustomerExcelDto
import com.example.demo.customer.Customer
import com.example.demo.customer.CustomerRepository
import com.example.demo.excel.sheet.SimpleExcelFile
import org.springframework.web.bind.annotation.*
import org.springframework.web.multipart.MultipartFile
import javax.servlet.http.HttpServletResponse

@RequestMapping(value = ["/api/excel"])
@RestController
class ExcelApi(private val customerRepository: CustomerRepository) {

    @GetMapping("/customers.xlsx")
    fun customerReport(httpServletResponse: HttpServletResponse) {
        val customers = customerRepository.findAll().toDto()

        val excel = SimpleExcelFile(customers, CustomerExcelDto::class.java)
        excel.write(httpServletResponse.outputStream)

        httpServletResponse.contentType = "ms-vnd/excel"
        httpServletResponse.setHeader("Content-Disposition", "attachment;filename=test.xls")
    }

    @PostMapping("/read")
    fun readExcel(@RequestParam("file") multipartFile: MultipartFile) {
//        val sheet = excelService.readExcel(multipartFile)
//        val sheet2 = excelService.readExcelTest(multipartFile)
//        println(sheet)
//        println(sheet2)
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
        price3 = 1234567
    )
}