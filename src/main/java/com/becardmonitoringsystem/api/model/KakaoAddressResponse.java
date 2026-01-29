package com.becardmonitoringsystem.api.model;

import java.util.List;

public record KakaoAddressResponse(
        List<Document> documents
) {
    public record Document(
            String x, // 경도 (longitude)
            String y  // 위도 (latitude)
    ) {}
}
