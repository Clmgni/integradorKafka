package com.clmgni.api.exception

import org.springframework.http.HttpHeaders
import org.springframework.http.HttpStatus
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.MethodArgumentNotValidException
import org.springframework.web.bind.annotation.ControllerAdvice
import org.springframework.web.context.request.WebRequest
import org.springframework.web.servlet.mvc.method.annotation.ResponseEntityExceptionHandler
import java.time.LocalDate
import java.util.stream.Collectors

@ControllerAdvice
class GenericException : ResponseEntityExceptionHandler() {

    override fun handleMethodArgumentNotValid(ex: MethodArgumentNotValidException, headers: HttpHeaders, status: HttpStatus, request: WebRequest): ResponseEntity<Any> {
        val body = HashMap<String, Any>().apply {
            put("timestamp", LocalDate.now())
            put("status", status.value())
            put("errors",  ex.bindingResult.fieldErrors.stream().map { x-> x.defaultMessage }.collect(Collectors.toList()))
        }

        return ResponseEntity(body, headers, status)
    }

}