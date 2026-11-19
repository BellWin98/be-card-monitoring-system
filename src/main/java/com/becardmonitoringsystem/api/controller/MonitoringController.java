package com.becardmonitoringsystem.api.controller;

import com.becardmonitoringsystem.api.model.AnalysisJob;
import com.becardmonitoringsystem.api.model.JobStatusResponse;
import com.becardmonitoringsystem.api.service.AnalysisJobService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Map;

@RestController
@RequestMapping("/api/monitor")
public class MonitoringController {

    private final AnalysisJobService analysisJobService;

    public MonitoringController(AnalysisJobService analysisJobService) {
        this.analysisJobService = analysisJobService;
    }

    /**
     * 분석 작업을 등록하고 즉시 jobId를 반환한다. (분석은 백그라운드에서 진행)
     */
    @PostMapping("/analyze")
    public ResponseEntity<JobStatusResponse> analyzeCardUsage(@RequestParam("file") MultipartFile file) throws IOException {
        AnalysisJob job = analysisJobService.submit(file);
        return ResponseEntity.status(HttpStatus.ACCEPTED).body(JobStatusResponse.from(job));
    }

    /**
     * 분석 작업 진행 상태 조회
     */
    @GetMapping("/jobs/{jobId}")
    public ResponseEntity<?> getJobStatus(@PathVariable String jobId) {
        return analysisJobService.find(jobId)
                .<ResponseEntity<?>>map(job -> ResponseEntity.ok(JobStatusResponse.from(job)))
                .orElseGet(MonitoringController::jobNotFound);
    }

    /**
     * 완료된 분석 결과 엑셀 다운로드
     */
    @GetMapping("/jobs/{jobId}/result")
    public ResponseEntity<?> downloadResult(@PathVariable String jobId) {
        return analysisJobService.find(jobId)
                .<ResponseEntity<?>>map(job -> {
                    if (job.getStatus() != AnalysisJob.Status.COMPLETED) {
                        return ResponseEntity.status(HttpStatus.CONFLICT)
                                .body(Map.of("message", "분석이 아직 완료되지 않았습니다."));
                    }
                    return ResponseEntity.ok()
                            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=analysis_result.xlsx")
                            .contentType(MediaType.APPLICATION_OCTET_STREAM)
                            .body(job.getResult());
                })
                .orElseGet(MonitoringController::jobNotFound);
    }

    private static ResponseEntity<?> jobNotFound() {
        return ResponseEntity.status(HttpStatus.NOT_FOUND)
                .body(Map.of("message", "분석 작업을 찾을 수 없습니다. 만료되었거나 서버가 재시작되었을 수 있습니다."));
    }
}
