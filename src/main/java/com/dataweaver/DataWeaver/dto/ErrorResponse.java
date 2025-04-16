package com.dataweaver.DataWeaver.dto;

import java.time.LocalDateTime;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ErrorResponse {

    private String message;
    private LocalDateTime timestamp;
    private String details;

    public ErrorResponse(String message, String details) {
        this.message = message;
        this.timestamp = LocalDateTime.now();
        this.details = details;
    }

}
