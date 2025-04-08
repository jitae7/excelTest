package fastexcelTest;

import common.Config;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Sheet;
import org.dhatim.fastexcel.reader.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class FastMultiTableExcelReader {
    public static void main(String[] args) {
        fastExcelTest(Config.FILE_PATH, Config.OUTPUT_PATH_FAST);
    }

    public static void fastExcelTest(String filePath, String outputFilePath) {
        long startTime = System.nanoTime();
        AtomicInteger totalCellCount = new AtomicInteger(0);
        StringBuilder outputContent = new StringBuilder();

        File file = new File(filePath);
        if (!file.exists()) {
            System.out.println("❌ 파일을 찾을 수 없습니다: " + filePath);
            return;
        }
        System.out.println("✅ 파일을 찾았습니다: " + filePath);

        try (FileInputStream fis = new FileInputStream(file);
             ReadableWorkbook wb = new ReadableWorkbook(fis)) {

            List<Sheet> sheets = StreamSupport.stream(wb.getSheets().spliterator(), false).collect(Collectors.toList());
            System.out.println("📊 총 " + sheets.size() + "개의 시트를 찾았습니다.");

            for (Sheet sheet : sheets) {
                outputContent.append("📄 [시트 이름]: ").append(sheet.getName()).append("\n");
                System.out.println("📄 현재 처리 중인 시트: " + sheet.getName());

                try (Stream<Row> rowStream = sheet.openStream()) {
                    rowStream.forEach(row -> {
                        StringBuilder rowData = new StringBuilder();
                        //1. FastExcel에서는 기본적으로 빈 셀을 null로 처리 
//                        row.stream().forEach(cell -> { // getCells() 대신 stream() 사용
//                            if (cell != null) { // null 체크 추가
//                                totalCellCount.incrementAndGet();
//                                rowData.append(cell.getRawValue()).append(" | ");
//                            } else {
//                                rowData.append("EMPTY | "); // 빈 셀 처리
//                            }
//                        });
                        //2. FastExcel에서도 빈 셀을 "EMPTY"로 출력하기 위해 null 체크 및 기본값 할당
                        row.stream().forEach(cell -> { 
                            // 셀 값이 null이면 "EMPTY" 출력
                            String cellValue = (cell == null || cell.getRawValue() == null) ? "EMPTY" : cell.getRawValue();
                            totalCellCount.incrementAndGet();
                            rowData.append(cellValue).append(" | ");
                        });
                        outputContent.append(rowData.toString()).append("\n");
                    });
                }
                outputContent.append("----------------------\n");
                System.out.println("✔ 시트 [" + sheet.getName() + "] 처리 완료");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        long endTime = System.nanoTime();
        long executionTimeMs = (endTime - startTime) / 1_000_000; // 실행 시간 (ms 단위)

        // 실행 시간을 파일 맨 위에 추가
        StringBuilder finalOutput = new StringBuilder();
        finalOutput.append("Execution Time: ").append(executionTimeMs).append(" ms\n");
        finalOutput.append(outputContent); // 기존 내용 추가
        finalOutput.append("총 셀 수: ").append(totalCellCount.get()).append("\n");
        finalOutput.append("FastExcel Execution Time: ").append(executionTimeMs).append(" ms\n");

        try (FileWriter writer = new FileWriter(outputFilePath)) {
            writer.write(finalOutput.toString());
            System.out.println("✅ 결과가 " + outputFilePath + " 파일에 저장되었습니다.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}