package poiExcelTest;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import common.Config;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.atomic.AtomicInteger;

public class PoiExcelTest {
    public static void main(String[] args) throws InvalidFormatException {
        poiTest(Config.FILE_PATH);
    }

    public static void poiTest(String filePath) throws InvalidFormatException {
        long startTime = System.nanoTime();
        AtomicInteger totalCellCount = new AtomicInteger(0);

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) { // 모든 시트 읽기
                Sheet sheet = workbook.getSheetAt(i);
                System.out.println("📄 [시트 이름]: " + sheet.getSheetName());

                for (Row row : sheet) {
                    for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                        Cell cell = row.getCell(cellNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // 빈 셀 포함
                        totalCellCount.incrementAndGet();
                        System.out.println(totalCellCount.get() + "##" + getCellValue(cell));
                    }
                }
                System.out.println("----------------------");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        long endTime = System.nanoTime();
        System.out.println("총 셀 수: " + totalCellCount.get());
        System.out.println("POI Execution Time: " + (endTime - startTime) / 1_000_000 + " ms");
    }

    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "EMPTY";
        }
    }
}

