package common; // 공통 패키지에 위치

public class Config {
    public static final String BASE_DIRECTORY = "C:\\Users\\jamong\\Desktop\\odor\\통계현황\\FastExcel vs. Apache POI\\"; // 공통 경로 지정
    public static final String FILE_NAME = "aa3.xlsx"; // 입력 파일 이름
    public static final String OUTPUT_FILE_NAME_FAST = "output_fast_FormulaEvaluator.txt"; // 출력 파일 이름
    public static final String OUTPUT_FILE_NAME_POI_XSSFWORKBOOK = "output_poi_xssfworkbook_FormulaEvaluator.txt"; // 출력 파일 이름
    public static final String OUTPUT_FILE_NAME_POI_XSSFREADER = "output_poi_xssfreader_FormulaEvaluator.txt"; // 출력 파일 이름

    public static final String FILE_PATH = BASE_DIRECTORY + FILE_NAME; // 전체 파일 경로
    public static final String OUTPUT_PATH_FAST = BASE_DIRECTORY + OUTPUT_FILE_NAME_FAST; // FastExcel 전체 출력 경로
    public static final String OUTPUT_PATH_POI_XSSFWORKBOOK = BASE_DIRECTORY + OUTPUT_FILE_NAME_POI_XSSFWORKBOOK; // Poi 전체 출력 경로
    public static final String OUTPUT_PATH_POI_XSSFREADER = BASE_DIRECTORY + OUTPUT_FILE_NAME_POI_XSSFREADER; // Poi 전체 출력 경로 (SAX 기반 API 사용)
}
