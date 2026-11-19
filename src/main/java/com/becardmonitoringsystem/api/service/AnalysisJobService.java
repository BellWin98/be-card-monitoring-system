package com.becardmonitoringsystem.api.service;

import com.becardmonitoringsystem.api.model.AnalysisJob;
import jakarta.annotation.PreDestroy;
import lombok.extern.slf4j.Slf4j;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.Instant;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * 엑셀 분석을 HTTP 요청과 분리해 백그라운드에서 실행한다.
 * 클라이언트는 jobId로 진행 상태를 조회하고, 완료되면 결과 파일을 별도로 내려받는다.
 */
@Service
@Slf4j
public class AnalysisJobService {

    /** 완료/실패한 작업 결과를 보관하는 시간 */
    private static final Duration RESULT_TTL = Duration.ofHours(1);

    private final ExcelAnalysisService excelAnalysisService;
    private final Map<String, AnalysisJob> jobs = new ConcurrentHashMap<>();

    // 카카오 API Rate Limit 보호를 위해 작업은 한 번에 하나씩 순차 처리한다.
    private final ExecutorService executor = Executors.newSingleThreadExecutor(r -> {
        Thread t = new Thread(r, "excel-analysis-worker");
        t.setDaemon(true);
        return t;
    });

    public AnalysisJobService(ExcelAnalysisService excelAnalysisService) {
        this.excelAnalysisService = excelAnalysisService;
    }

    public AnalysisJob submit(MultipartFile file) throws IOException {
        // 업로드 임시파일은 요청 종료 시 삭제되므로 바이트를 미리 복사해 둔다.
        byte[] content = file.getBytes();

        AnalysisJob job = new AnalysisJob(UUID.randomUUID().toString());
        jobs.put(job.getId(), job);
        executor.submit(() -> run(job, content));

        log.info("분석 작업 등록: {} ({})", job.getId(), file.getOriginalFilename());
        return job;
    }

    public Optional<AnalysisJob> find(String jobId) {
        return Optional.ofNullable(jobs.get(jobId));
    }

    private void run(AnalysisJob job, byte[] content) {
        job.start();
        try {
            byte[] result = excelAnalysisService.processExcel(new ByteArrayInputStream(content), job::updateProgress);
            job.complete(result);
            log.info("분석 작업 완료: {} ({}건)", job.getId(), job.getTotal());
        } catch (Exception e) {
            log.error("분석 작업 실패: {}", job.getId(), e);
            job.fail(e.getMessage() != null ? e.getMessage() : "분석 중 오류가 발생했습니다.");
        }
    }

    @Scheduled(fixedDelay = 10 * 60 * 1000)
    public void removeExpiredJobs() {
        Instant threshold = Instant.now().minus(RESULT_TTL);
        jobs.values().removeIf(job -> job.getFinishedAt() != null && job.getFinishedAt().isBefore(threshold));
    }

    @PreDestroy
    public void shutdown() {
        executor.shutdownNow();
    }
}
