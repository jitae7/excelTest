package poiExcelTest;

import common.Config;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

public class PoiMultiTableExcelReader {
    public static void main(String[] args) throws InvalidFormatException {
        poiTest(Config.FILE_PATH, Config.OUTPUT_PATH_POI_XSSFWORKBOOK);
    }

    public static void poiTest(String filePath, String outputFilePath) throws InvalidFormatException {
        File file = new File(filePath);
        if (!file.exists()) {
            System.out.println("❌ 파일을 찾을 수 없습니다: " + filePath);
            return;
        }
        System.out.println("✅ 파일을 찾았습니다: " + filePath);

        // 출력 파일 처리: 파일이 있으면 덮어쓰고, 없으면 새로 생성
        File outputFile = new File(outputFilePath);
        if (outputFile.exists()) {
            System.out.println("✅ 출력 파일이 존재합니다. 덮어씁니다: " + outputFilePath);
        } else {
            System.out.println("✅ 출력 파일이 존재하지 않습니다. 새로 생성합니다: " + outputFilePath);
        }

        long startTime = System.nanoTime();
        AtomicInteger totalCellCount = new AtomicInteger(0);

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis);
             // FileWriter의 두 번째 인자를 false로 전달하여 덮어쓰기 모드로 생성
             BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile, false))) {

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator(); // ✅ 수식 평가기 한 번만 생성

            int sheetCount = workbook.getNumberOfSheets();
            System.out.println("📊 총 " + sheetCount + "개의 시트를 찾았습니다.");
            writer.write("Execution Time: (calculating...) ms\n"); // ✅ 실행 시간 자리 확보
            writer.newLine();

            StringBuilder outputContent = new StringBuilder();
            
            for (int i = 0; i < sheetCount; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                outputContent.append("📄 [시트 이름]: ").append(sheet.getSheetName()).append("\n");
                System.out.println("📄 현재 처리 중인 시트: " + sheet.getSheetName());

                List<List<String>> tables = new ArrayList<>();
                List<String> currentTable = new ArrayList<>();
                int rowCount = 0;

                for (Row row : sheet) {
                    rowCount++;
                    boolean isEmptyRow = isRowEmpty(row);
                    if (isEmptyRow && !currentTable.isEmpty()) {
                        tables.add(new ArrayList<>(currentTable));
                        currentTable.clear();
                    } else if (!isEmptyRow) {
                        StringBuilder rowData = new StringBuilder();
                        for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                            Cell cell = row.getCell(cellNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            totalCellCount.incrementAndGet();
                            rowData.append(getCellValue(cell, evaluator)).append(" | ");
                        }
                        currentTable.add(rowData.toString());
                    }
                }

                if (!currentTable.isEmpty()) {
                    tables.add(currentTable);
                }

                int tableIndex = 1;
                for (List<String> table : tables) {
                    outputContent.append("===== Table ").append(tableIndex++).append(" =====\n");
                    for (String row : table) {
                        outputContent.append(row).append("\n");
                    }
                    outputContent.append("----------------------\n");
                }

                System.out.println("✔ 시트 [" + sheet.getSheetName() + "] 처리 완료 (행 개수: " + rowCount + ")");

                // ✅ 중간 결과를 파일에 저장 (메모리 초과 방지)
                writer.write(outputContent.toString());
                outputContent.setLength(0); // StringBuilder 초기화
            }

            long endTime = System.nanoTime();
            long executionTimeMs = (endTime - startTime) / 1_000_000; // 실행 시간 (ms 단위)

            // ✅ 실행 시간 파일에 추가
            writer.write("총 셀 수: " + totalCellCount.get() + "\n");
            writer.write("POI Execution Time: " + executionTimeMs + " ms\n");
            
            System.out.println("✅ 실행 시간: " + executionTimeMs + " ms");
            System.out.println("✅ 결과가 " + outputFilePath + " 파일에 저장되었습니다.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (Cell cell : row) {
            if (cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }

    private static String getCellValue(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return "EMPTY";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return evaluateFormula(cell, evaluator);
            default:
                return "EMPTY";
        }
    }

    private static String evaluateFormula(Cell cell, FormulaEvaluator evaluator) {
        try {
            CellValue cellValue = evaluator.evaluate(cell);
            if (cellValue == null) return "EMPTY";

            switch (cellValue.getCellType()) {
                case STRING:
                    return cellValue.getStringValue();
                case NUMERIC:
                    return String.valueOf(cellValue.getNumberValue());
                case BOOLEAN:
                    return String.valueOf(cellValue.getBooleanValue());
                default:
                    return "EMPTY";
            }
        } catch (Exception e) {
            return "FORMULA_ERROR";
        }
    }
}
