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

@RestController
@RequestMapping("/data-weaver")
public class DataWeaverController {

    private DataWeaverService dataWeaverService;

    public DataWeaverController(DataWeaverService dataWeaverService) {
        this.dataWeaverService = dataWeaverService;
    }

    @PostMapping("/generate-excel")
    public ResponseEntity<byte[]> generateExcel(@RequestParam("file") MultipartFile file, @RequestParam("month") int month, @RequestParam("year") int year) throws IOException {
        byte[] outputBytes = dataWeaverService.generateExcel(file, month, year);


        return ResponseEntity.ok()  
                .header("Content-Disposition", "attachment; filename=\"output.xlsx\"")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(outputBytes);
    }


    @GetMapping("/health-check")
    public ResponseEntity<String> healthCheckup() {
        return ResponseEntity.ok()
                .body("Data-Weaver Service is up and healthy");
    }

}
