package com.becardmonitoringsystem.api.service;

import com.becardmonitoringsystem.api.model.KakaoAddressResponse;
import com.becardmonitoringsystem.api.model.KakaoAddressResponse.Document;
import io.netty.handler.ssl.SslContext;
import io.netty.handler.ssl.SslContextBuilder;
import io.netty.handler.ssl.util.InsecureTrustManagerFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.client.reactive.ReactorClientHttpConnector;
import org.springframework.stereotype.Service;
import org.springframework.web.reactive.function.client.WebClient;
import reactor.core.publisher.Mono;
import reactor.netty.http.client.HttpClient;

import javax.net.ssl.SSLException;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

@Slf4j
@Service
public class KakaoMapService {

    @Value("${kakao.api.key}")
    private String apiKey;

    @Value("${kakao.api.url}")
    private String apiUrl;

    private final WebClient webClient;

    /** 주소 → 좌표(Document) 캐시 (동일 주소 재요청 시 API 호출 생략) */
    private final Map<String, Document> addressCache = new ConcurrentHashMap<>();

    public KakaoMapService() throws SSLException {
        // SSL 검증을 무시하는 SslContext 생성
        SslContext sslContext = SslContextBuilder
                .forClient()
                .trustManager(InsecureTrustManagerFactory.INSTANCE)
                .build();

        // HttpClient에 SSL 설정 적용
        HttpClient httpClient = HttpClient.create()
                .secure(t -> t.sslContext(sslContext));

        // WebClient 빌드
        this.webClient = WebClient.builder()
                .clientConnector(new ReactorClientHttpConnector(httpClient))
                .build();
    }

    /**
     * 주소를 좌표로 변환 (요구사항: 1차 실패 시 2차 시도).
     * 동일 주소 재요청 시 캐시된 좌표를 반환하여 API 호출을 생략한다.
     */
    public Document getCoordinates(String addr1, String addr2) {
        // 1차: 주소1 + 주소2
        String fullAddress = (addr1 + " " + (addr2 != null ? addr2 : "")).trim();
        Document cached = addressCache.get(fullAddress);
        if (cached != null) {
            return cached;
        }

        Document result = fetchFromApi(fullAddress);
        if (result != null) {
            addressCache.put(fullAddress, result);
            return result;
        }

        // 2차 시도: 검색 결과가 없고 addr2가 존재했다면 addr1으로만 재시도
        if (addr2 != null && !addr2.isBlank()) {
            result = fetchFromApi(addr1);
        if (result != null) {
            addressCache.put(fullAddress, result);
        }
    }

        return result;
    }

    private Document fetchFromApi(String address) {
        return webClient.get()
                .uri(apiUrl + "?query=" + address) // 전체 URL을 직접 전달
                .header("Authorization", "KakaoAK " + apiKey)
                .retrieve()
                .bodyToMono(KakaoAddressResponse.class)
                .flatMap(response -> {
                    // 로그로 응답 데이터 확인
//                    log.info("API 응답 결과: {}", response);

                    if (response.documents() != null && !response.documents().isEmpty()) {
                        // 결과가 있으면 첫 번째 결과 반환
                        return Mono.just(response.documents().getFirst());
                    }
                    // 결과가 없으면 빈 객체가 아닌 Empty Mono 반환
                    return Mono.empty();
                })
                .block();
    }
}
