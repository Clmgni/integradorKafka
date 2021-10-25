package com.clmgni.api.kafka

import com.clmgni.api.config.KafkaConfig
import com.fasterxml.jackson.databind.JsonNode
import com.fasterxml.jackson.databind.ObjectMapper
import org.apache.kafka.clients.producer.ProducerRecord
import org.apache.kafka.clients.producer.RecordMetadata
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.kafka.core.KafkaTemplate
import org.springframework.stereotype.Service
import java.lang.Exception

@Service
class SendToKafka(val kafkaConfig : KafkaConfig)  {

    @Autowired
    private lateinit var kafkaTemplate : KafkaTemplate<String, String>

    private val producerJson =  kafkaConfig.producerJson()

    fun sendToKafkaJson(topic : String, obj : Any) {
        val node = ObjectMapper().valueToTree<JsonNode>(obj)
        val record : ProducerRecord<String, JsonNode> = ProducerRecord(topic, node)

        producerJson.send(record) { metadata: RecordMetadata?, exception: Exception? ->
            println(exception ?: metadata)
        }
    }


}