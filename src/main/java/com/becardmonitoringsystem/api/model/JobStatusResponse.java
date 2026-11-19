package com.becardmonitoringsystem.api.model;

public record JobStatusResponse(
        String jobId,
        AnalysisJob.Status status,
        int processed,
        int total,
        String message
) {
    public static JobStatusResponse from(AnalysisJob job) {
        return new JobStatusResponse(
                job.getId(),
                job.getStatus(),
                job.getProcessed(),
                job.getTotal(),
                job.getErrorMessage()
        );
    }
}
