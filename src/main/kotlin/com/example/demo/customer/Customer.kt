package com.example.demo.customer

import org.springframework.data.jpa.repository.JpaRepository
import java.time.Instant
import javax.persistence.*

@Entity
@Table(name = "customer")
class Customer(
    @Id
    @GeneratedValue
    var id: Long = 0,

    @Column(name = "name")
    var name: String,

    @Column(name = "address")
    var address: String,

    @Column(name = "age")
    var age: Int = 0,

    @Column(name = "status")
    var status: Status,

    @Column(name = "created_at")
    var createdAt: Instant
) {
    override fun toString(): String {
        return "Customer(id=$id, name='$name', address='$address', age=$age, status=$status, createdAt=$createdAt)"
    }
}

enum class Status {
    ACTIVE, INACTIVE
}

interface CustomerRepository : JpaRepository<Customer, Long>