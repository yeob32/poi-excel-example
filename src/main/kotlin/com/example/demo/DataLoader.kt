package com.example.demo

import com.example.demo.customer.Customer
import com.example.demo.customer.CustomerRepository
import com.example.demo.customer.Status
import org.springframework.boot.ApplicationRunner
import org.springframework.context.annotation.Bean
import org.springframework.stereotype.Component
import java.time.Instant
import java.time.temporal.ChronoUnit

@Component
class DataLoader(private val customerRepository: CustomerRepository) {

    @Bean
    fun applicationRunner(): ApplicationRunner = ApplicationRunner {
        customerRepository.saveAll(
            listOf(
                Customer(
                    name = "Jack Smith",
                    address = "Massachusetts",
                    age = 23,
                    createdAt = Instant.now(),
                    status = Status.ACTIVE
                ),
                Customer(
                    name = "Adam Johnson",
                    address = "New York",
                    age = 27,
                    createdAt = Instant.now().plus(-50, ChronoUnit.DAYS),
                    status = Status.ACTIVE
                ),
                Customer(
                    name = "Katherin Carter",
                    address = "Washington DC",
                    age = 26,
                    createdAt = Instant.now().plus(-20, ChronoUnit.DAYS),
                    status = Status.INACTIVE
                ),
                Customer(
                    name = "Jack London",
                    address = "Nevada",
                    age = 33,
                    createdAt = Instant.now().plus(30, ChronoUnit.DAYS),
                    status = Status.INACTIVE
                ),
                Customer(
                    name = "Jason Bourne",
                    address = "California",
                    age = 36,
                    createdAt = Instant.now().plus(10, ChronoUnit.DAYS),
                    status = Status.INACTIVE
                )
            )
        )
    }
}