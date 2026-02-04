package com.becardmonitoringsystem.api.service;

import com.becardmonitoringsystem.api.util.DistanceCalculator;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

@Service
@Slf4j
public class ExcelAnalysisService {

    private static final double DISTANCE = 2.0;
    private final KakaoMapService kakaoMapService;
    private final DistanceCalculator distanceCalculator; // 하버사인 공식 구현체

    public ExcelAnalysisService(KakaoMapService kakaoMapService, DistanceCalculator distanceCalculator) {
        this.kakaoMapService = kakaoMapService;
        this.distanceCalculator = distanceCalculator;
    }

    public byte[] processExcel(MultipartFile file) {
        try (InputStream is = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(is);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            Row header = sheet.getRow(0);

            int lastCol = header.getLastCellNum();
            header.createCell(lastCol).setCellValue("직선거리(km)");
            header.createCell(lastCol + 1).setCellValue("결과");

            // 데이터 행 반복 (헤더 제외)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // 1. 개별 행 처리 (여기서 block()이 발생해도 안전함)
                processRowDetail(row, lastCol);

                // 2. 인위적인 지연 (Rate Limit 적용)
                try {
                    Thread.sleep(100); // 0.1초 대기
                } catch (InterruptedException ie) {
                    Thread.currentThread().interrupt();
                }
            }

            workbook.write(bos);
            return bos.toByteArray();

        } catch (Exception e) {
            throw new RuntimeException("엑셀 처리 중 오류 발생: " + e.getMessage());
        }
    }

    /**
     * 개별 행에 대한 좌표 수집 및 거리 계산 로직
     */
    private void processRowDetail(Row row, int lastCol) {
        try {
            String storeAddr1 = getCellValue(row.getCell(13));
            String storeAddr2 = getCellValue(row.getCell(14));
            String homeAddr1 = getCellValue(row.getCell(45));
            String homeAddr2 = getCellValue(row.getCell(46));

            // 가맹점 주소: 둘 다 비어있으면 에러
            if (isBlank(storeAddr1) && isBlank(storeAddr2)) {
                row.createCell(lastCol + 1).setCellValue("가맹점주소 비어있음");
                return;
            }

            // 자택 주소: 둘 다 비어있으면 에러
            if (isBlank(homeAddr1) && isBlank(homeAddr2)) {
                row.createCell(lastCol + 1).setCellValue("자택주소 비어있음");
                return;
            }

            // 자택 좌표 수집
            var homeCoord = kakaoMapService.getCoordinates(homeAddr1, homeAddr2);
            if (homeCoord == null) {
                row.createCell(lastCol + 1).setCellValue("자택주소 검색결과 없음");
                return;
            }

            // 가맹점 좌표 수집
            var storeCoord = kakaoMapService.getCoordinates(storeAddr1, storeAddr2);
            if (storeCoord == null) {
                row.createCell(lastCol + 1).setCellValue("가맹점주소 검색결과 없음");
                return;
            }

            // 거리 계산
            double distance = distanceCalculator.calculate(
                    Double.parseDouble(storeCoord.y()), Double.parseDouble(storeCoord.x()),
                    Double.parseDouble(homeCoord.y()), Double.parseDouble(homeCoord.x())
            );

            // 결과 기록
            row.createCell(lastCol).setCellValue(Math.round(distance * 100) / 100.0);
            row.createCell(lastCol + 1).setCellValue(distance <= DISTANCE ? "2km 이내" : "2km 초과");

        } catch (Exception e) {
            row.createCell(lastCol + 1).setCellValue("계산 중 오류: " + e.getMessage());
        }
    }

    private boolean isBlank(String s) {
        return s == null || s.isBlank();
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            default -> "";
        };
    }
}
