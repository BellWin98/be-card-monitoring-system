package com.becardmonitoringsystem.api.util;

import org.springframework.stereotype.Component;

@Component
public class DistanceCalculator {

    // 지구 반지름 (단위: km)
    private static final double EARTH_RADIUS = 6371;

    /**
     * 두 지점 간의 직선거리를 계산합니다.
     * @param lat1 지점1 위도
     * @param lon1 지점1 경도
     * @param lat2 지점2 위도
     * @param lon2 지점2 경도
     * @return 직선거리 (km)
     */
    public double calculate(double lat1, double lon1, double lat2, double lon2) {
        // 위도와 경도를 라디안(Radian)으로 변환
        double dLat = Math.toRadians(lat2 - lat1);
        double dLon = Math.toRadians(lon2 - lon1);

        double a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
                Math.cos(Math.toRadians(lat1)) * Math.cos(Math.toRadians(lat2)) *
                        Math.sin(dLon / 2) * Math.sin(dLon / 2);

        double c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

        return EARTH_RADIUS * c;
    }
}
