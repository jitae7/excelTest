package fastexcelTest;

import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Sheet;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class FastExcelTest {
    public static void main(String[] args) {
        File file = new File("C:\\aa3.xlsx"); // 엑셀 파일 경로

        if (file.exists()) {
            System.out.println("✅ 파일이 존재합니다: " + file.getAbsolutePath());
        } else {
            System.out.println("❌ 파일을 찾을 수 없습니다: " + file.getAbsolutePath());
            return;
        }

        AtomicInteger totalCellCount = new AtomicInteger(0);

        try (FileInputStream fis = new FileInputStream(file);
             ReadableWorkbook wb = new ReadableWorkbook(fis)) {

            List<Sheet> sheets = StreamSupport.stream(wb.getSheets().spliterator(), false).collect(Collectors.toList());
            
            for (Sheet sheet : sheets) {
                System.out.println("📄 [시트 이름]: " + sheet.getName());
                
                try (Stream<Row> rowStream = sheet.openStream()) {
                    rowStream.forEach(row -> {
                        for (Cell cell : row) {
                            int currentCount = totalCellCount.incrementAndGet();
                            if (cell != null) {
                                System.out.println(currentCount + "##" + cell.getRawValue());
                            } else {
                                System.out.println(currentCount + "## EMPTY");
                            }
                        }
                    });
                }

                System.out.println("----------------------");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("총 셀 수: " + totalCellCount.get());
    }
}
