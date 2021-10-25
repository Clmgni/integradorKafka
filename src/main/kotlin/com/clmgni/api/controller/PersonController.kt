package com.example.api.controller

import com.clmgni.api.dto.Person
import com.clmgni.api.kafka.SendToKafka
import com.clmgni.api.utils.Util
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.beans.factory.annotation.Value
import org.springframework.http.HttpStatus
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.*
import javax.validation.Valid

@RestController
@RequestMapping("/person")
class PersonController(val sendToKafka: SendToKafka) {

    @Value("\${topic-person}")
    lateinit var topic : String

    @PostMapping("/post}")
    fun post(@Valid @RequestBody person : Person) : ResponseEntity<Any> {
        person.cpf = Util.removeSpecialCaracterFromString(person.cpf)
        return try {
            sendToKafka.sendToKafkaJson(topic, person)
            ResponseEntity.ok().build()
        } catch (e : Exception) {
            ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Problema ao enviar mensagem")
        }
    }
}