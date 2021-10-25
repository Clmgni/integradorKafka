package com.clmgni.api.dto

import javax.validation.constraints.*

data class Person(
    @field:NotBlank(message = "\${name.notblank}")
    val name : String,

    @field:NotBlank(message = "\${cpf.notblank}")
    var cpf : String,

    val idade : Int
) {
    override fun toString() : String {
        return "{\"name\": \"$name\", \"cpf\":\"$cpf\", \"idade\": \"$idade\"}"
    }

}