package com.dataweaver.DataWeaver.controller;


import java.io.IOException;

import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.dataweaver.DataWeaver.service.DataWeaverService;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;

@Tag(name = "DataWeaver Controller", description = "Handles health checkup and generate-excel end points")
@RestController
@RequestMapping("/data-weaver")
public class DataWeaverController {

    private DataWeaverService dataWeaverService;

    public DataWeaverController(DataWeaverService dataWeaverService) {
        this.dataWeaverService = dataWeaverService;
    }

    @Operation(summary = "Returns back an excel sheet as response", description = "Returns back a polished excel sheet as response by seperating our each person's monthly tasks")
    @PostMapping("/generate-excel")
    public ResponseEntity<byte[]> generateExcel(@RequestParam("file") MultipartFile file) throws IOException {
        byte[] outputBytes = dataWeaverService.generateExcel(file);


        return ResponseEntity.ok()  
                .header("Content-Disposition", "attachment; filename=\"output.xlsx\"")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(outputBytes);
    }


    @Operation(summary = "Health check end point", description = "Checks whether DataWeaver application is up and healthy or not")
    @GetMapping("/health-check")
    public ResponseEntity<String> healthCheckup() {
        return ResponseEntity.ok()
                .body("Data-Weaver Service is up and healthy");
    }

}
