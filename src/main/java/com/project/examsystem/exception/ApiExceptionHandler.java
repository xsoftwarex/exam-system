package com.project.examsystem.exception;


import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

@RestControllerAdvice
public class ApiExceptionHandler {
    @ExceptionHandler(FileException.class)
    public ResponseEntity<ApiError> handleFileException(Exception exception) {
        List<String> errors = new ArrayList<>();
        errors.add(exception.getMessage());
        ApiError apiError = new ApiError(HttpStatus.INTERNAL_SERVER_ERROR, errors, LocalDateTime.now());
        return new ResponseEntity<>(apiError, apiError.getStatus());
    }
}

