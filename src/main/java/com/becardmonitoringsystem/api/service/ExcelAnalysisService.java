package com.becardmonitoringsystem.api.service;

import com.becardmonitoringsystem.api.util.DistanceCalculator;
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

    /*public byte[] processExcel(MultipartFile file) {
        try (InputStream is = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(is);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            Row header = sheet.getRow(0);

            int lastCol = header.getLastCellNum();
            header.createCell(lastCol).setCellValue("직선거리(km)");
            header.createCell(lastCol + 1).setCellValue("결과");

            // 1. 처리할 행의 인덱스 리스트 생성 (1번 행부터 마지막 행까지)
            var rowIndices = IntStream.rangeClosed(1, sheet.getLastRowNum()).boxed().toList();

            // 2. Flux를 사용하여 Rate Limit 적용
            Flux.fromIterable(rowIndices)
                    .delayElements(Duration.ofMillis(100)) // 0.1초 간격 (초당 최대 10건 호출)
                    .doOnNext(i -> {
                        Row row = sheet.getRow(i);
                        if (row != null) {
                            processRowDetail(row, lastCol); // 개별 행 처리 로직 분리
                        }
                    })
                    .blockLast(); // 모든 데이터가 처리될 때까지 대기

            workbook.write(bos);
            return bos.toByteArray();

        } catch (Exception e) {
            throw new RuntimeException("엑셀 처리 중 오류 발생: " + e.getMessage());
        }
    }*/

    /**
     * 개별 행에 대한 좌표 수집 및 거리 계산 로직
     */
    private void processRowDetail(Row row, int lastCol) {
        try {
            String storeAddr1 = getCellValue(row.getCell(13));
            String storeAddr2 = getCellValue(row.getCell(14));
            String homeAddr = getCellValue(row.getCell(45));

            // 자택 좌표 수집
            var homeCoord = kakaoMapService.getCoordinates(homeAddr, "");
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


    /*public byte[] processExcel(MultipartFile file) {
        try (InputStream is = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(is);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            Row header = sheet.getRow(0);

            // 결과 컬럼 추가 (비어있는 마지막 셀 뒤에 추가)
            int lastCol = header.getLastCellNum();
            header.createCell(lastCol).setCellValue("직선거리(km)");
            header.createCell(lastCol + 1).setCellValue("결과");

            // 데이터 행 반복 (헤더 제외)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // 엑셀 내 인덱스(예시): 가맹점주소1(0), 가맹점주소2(1), 자택주소(2)
                String storeAddr1 = getCellValue(row.getCell(13));
                String storeAddr2 = getCellValue(row.getCell(14));
                String homeAddr = getCellValue(row.getCell(45));

                // 1. 좌표 수집 (카카오 API)

                var homeCoord = kakaoMapService.getCoordinates(homeAddr, "");
                if (homeCoord == null) {
                    row.createCell(lastCol + 1).setCellValue("자택주소 검색결과 없음");
                    continue;
                }

                var storeCoord = kakaoMapService.getCoordinates(storeAddr1, storeAddr2);
                if (storeCoord == null) {
                    row.createCell(lastCol + 1).setCellValue("가맹점주소 검색결과 없음");
                    continue;
                }

                // 2. 거리 계산
                double distance = distanceCalculator.calculate(
                        Double.parseDouble(storeCoord.y()), Double.parseDouble(storeCoord.x()),
                        Double.parseDouble(homeCoord.y()), Double.parseDouble(homeCoord.x())
                );

                // 3. 결과 기록
                row.createCell(lastCol).setCellValue(Math.round(distance * 100) / 100.0);
                row.createCell(lastCol + 1).setCellValue(distance <= DISTANCE ? DISTANCE + "이내" : DISTANCE + "초과");
            }

            workbook.write(bos);
            return bos.toByteArray();

        } catch (Exception e) {
            throw new RuntimeException("엑셀 처리 중 오류 발생: " + e.getMessage());
        }
    }*/

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            default -> "";
        };
    }
}
