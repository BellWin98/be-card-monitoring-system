package com.becardmonitoringsystem.api.controller;

import com.becardmonitoringsystem.api.service.ExcelAnalysisService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/monitor")
public class MonitoringController {

    private final ExcelAnalysisService excelAnalysisService;

    public MonitoringController(ExcelAnalysisService excelAnalysisService) {
        this.excelAnalysisService = excelAnalysisService;
    }

    @PostMapping("/analyze")
    public ResponseEntity<byte[]> analyzeCardUsage(@RequestParam("file") MultipartFile file) {
        byte[] excelContent = excelAnalysisService.processExcel(file);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=analysis_result.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(excelContent);
    }
}
