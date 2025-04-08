package poiExcelTest;

import common.Config;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.concurrent.atomic.AtomicInteger;

public class PoiMultiTableExcelReaderSAX {

    public static void main(String[] args) throws Exception {
        saxTest(Config.FILE_PATH, Config.OUTPUT_PATH_POI_XSSFREADER);
    }

    public static void saxTest(String filePath, String outputFilePath) throws Exception {
        File file = new File(filePath);
        if (!file.exists()) {
            System.out.println("❌ 파일을 찾을 수 없습니다: " + filePath);
            return;
        }
        System.out.println("✅ 파일을 찾았습니다: " + filePath);

        // 출력 파일 처리: 파일이 있으면 덮어쓰기, 없으면 새로 생성
        File outputFile = new File(outputFilePath);
        if (outputFile.exists()) {
            System.out.println("✅ 출력 파일이 존재합니다. 덮어씁니다: " + outputFilePath);
        } else {
            System.out.println("✅ 출력 파일이 존재하지 않습니다. 새로 생성합니다: " + outputFilePath);
        }

        long startTime = System.nanoTime();
        AtomicInteger totalCellCount = new AtomicInteger(0);

        // OPCPackage를 읽기 전용 모드로 엽니다.
        try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ);
             BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile, false))) {

            // XLSX 스트리밍 읽기를 위한 준비: 스타일, 공유 문자열 테이블
            XSSFReader xssfReader = new XSSFReader(pkg);
            StylesTable styles = xssfReader.getStylesTable();
            ReadOnlySharedStringsTable sharedStrings = new ReadOnlySharedStringsTable(pkg);

            // 시트 단위로 순회
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            int sheetIndex = 0;
            StringBuilder outputContent = new StringBuilder();
            writer.write("Execution Time: (calculating...) ms\n\n");

            while (iter.hasNext()) {
                InputStream sheetStream = iter.next();
                String sheetName = iter.getSheetName();
                sheetIndex++;
                outputContent.append("📄 [시트 이름]: ").append(sheetName).append("\n");
                System.out.println("📄 현재 처리 중인 시트: " + sheetName);

                // SAX 파서를 생성하고 커스텀 핸들러를 설정
                XMLReader parser = XMLReaderFactory.createXMLReader();
                SheetHandler handler = new SheetHandler(styles, sharedStrings, totalCellCount);
                parser.setContentHandler(handler);

                // 시트 스트림을 파싱
                InputSource sheetSource = new InputSource(sheetStream);
                parser.parse(sheetSource);
                sheetStream.close();

                // 핸들러에서 생성한 시트의 출력 내용을 가져옴
                outputContent.append(handler.getOutput());
                outputContent.append("\n");

                // 중간 결과를 파일에 기록하여 메모리 사용을 줄임
                writer.write(outputContent.toString());
                writer.flush();
                outputContent.setLength(0);
            }

            long endTime = System.nanoTime();
            long executionTimeMs = (endTime - startTime) / 1_000_000;
            writer.write("총 셀 수: " + totalCellCount.get() + "\n");
            writer.write("SAX Execution Time: " + executionTimeMs + " ms\n");
            System.out.println("✅ 실행 시간: " + executionTimeMs + " ms");
            System.out.println("✅ 결과가 " + outputFilePath + " 파일에 저장되었습니다.");
        }
    }

    /**
     * SAX 기반으로 시트의 XML 데이터를 읽어 각 행/셀의 값을 처리하는 핸들러.
     * 빈 행이 감지되면 현재까지의 데이터를 하나의 테이블로 구분하여 출력 문자열에 추가합니다.
     */
    private static class SheetHandler extends DefaultHandler {
        private final StylesTable stylesTable;
        private final ReadOnlySharedStringsTable sharedStringsTable;
        private final AtomicInteger totalCellCount;
        private final StringBuilder sheetOutput = new StringBuilder();
        private final StringBuilder currentTable = new StringBuilder();

        private StringBuilder value = new StringBuilder();
        private boolean vIsOpen;
        private String cellType;
        private boolean rowHasData;
        private StringBuilder rowContent = new StringBuilder();
        private int tableIndex = 1;

        public SheetHandler(StylesTable styles, ReadOnlySharedStringsTable sst, AtomicInteger totalCellCount) {
            this.stylesTable = styles;
            this.sharedStringsTable = sst;
            this.totalCellCount = totalCellCount;
        }

        public String getOutput() {
            // 마지막 테이블이 남아있으면 출력
            if (currentTable.length() > 0) {
                sheetOutput.append("===== Table ").append(tableIndex++).append(" =====\n");
                sheetOutput.append(currentTable);
                sheetOutput.append("----------------------\n");
                currentTable.setLength(0);
            }
            return sheetOutput.toString();
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
            if ("row".equals(qName)) {
                // 새 행 시작 시 초기화
                rowContent.setLength(0);
                rowHasData = false;
            } else if ("c".equals(qName)) {
                // 셀 시작: 셀 타입을 확인하고 값 버퍼 초기화
                totalCellCount.incrementAndGet();
                cellType = attributes.getValue("t");
                value.setLength(0);
            } else if ("v".equals(qName) || "is".equals(qName)) {
                vIsOpen = true;
                value.setLength(0);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            if (vIsOpen) {
                value.append(ch, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) throws SAXException {
            if ("v".equals(qName) || "t".equals(qName)) {
                vIsOpen = false;
            } else if ("c".equals(qName)) {
                // 셀 종료 시, 셀 값을 가져와 행 문자열에 추가
                String cellValue = getCellValue(value.toString(), cellType);
                rowContent.append(cellValue).append(" | ");
                if (!cellValue.trim().isEmpty() && !"EMPTY".equals(cellValue)) {
                    rowHasData = true;
                }
            } else if ("row".equals(qName)) {
                // 행 종료: 행이 비어있으면 테이블 경계로 간주
                if (rowHasData) {
                    currentTable.append(rowContent.toString()).append("\n");
                } else {
                    if (currentTable.length() > 0) {
                        sheetOutput.append("===== Table ").append(tableIndex++).append(" =====\n");
                        sheetOutput.append(currentTable);
                        sheetOutput.append("----------------------\n");
                        currentTable.setLength(0);
                    }
                }
            }
        }

        /**
         * 셀 타입에 따라 값을 해석합니다.
         */
        private String getCellValue(String valueStr, String cellType) {
            if ("s".equals(cellType)) { // 공유 문자열
                try {
                    int idx = Integer.parseInt(valueStr);
                    return sharedStringsTable.getEntryAt(idx).toString();
                } catch (Exception e) {
                    return valueStr;
                }
            }
            if (valueStr == null || valueStr.isEmpty()) {
                return "EMPTY";
            }
            return valueStr;
        }
    }
}
