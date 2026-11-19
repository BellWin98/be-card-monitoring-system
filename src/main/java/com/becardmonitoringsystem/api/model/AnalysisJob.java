package com.becardmonitoringsystem.api.model;

import java.time.Instant;

/**
 * 비동기 엑셀 분석 작업의 상태를 보관한다.
 * 분석 스레드가 쓰고 HTTP 요청 스레드가 읽으므로 필드는 volatile로 선언한다.
 */
public class AnalysisJob {

    public enum Status { PENDING, RUNNING, COMPLETED, FAILED }

    private final String id;
    private volatile Status status = Status.PENDING;
    private volatile int processed;
    private volatile int total;
    private volatile byte[] result;
    private volatile String errorMessage;
    private volatile Instant finishedAt;

    public AnalysisJob(String id) {
        this.id = id;
    }

    public void start() {
        this.status = Status.RUNNING;
    }

    public void updateProgress(int processed, int total) {
        this.processed = processed;
        this.total = total;
    }

    public void complete(byte[] result) {
        this.result = result;
        this.finishedAt = Instant.now();
        this.status = Status.COMPLETED;
    }

    public void fail(String errorMessage) {
        this.errorMessage = errorMessage;
        this.finishedAt = Instant.now();
        this.status = Status.FAILED;
    }

    public String getId() { return id; }
    public Status getStatus() { return status; }
    public int getProcessed() { return processed; }
    public int getTotal() { return total; }
    public byte[] getResult() { return result; }
    public String getErrorMessage() { return errorMessage; }
    public Instant getFinishedAt() { return finishedAt; }
}
