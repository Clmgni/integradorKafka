package com.clmgni.api

import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication

@SpringBootApplication
class KafkaJsonKotlinProducerApplication

fun main(args: Array<String>) {
    runApplication<KafkaJsonKotlinProducerApplication>(*args)
}