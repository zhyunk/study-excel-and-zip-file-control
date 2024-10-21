package excel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class ExcelService {
    private final String FILE_DIR_PATH = "/home/zhyun/my-projects/study-excel-and-zip-file-control/excel/%s";
    private final String NEW_FILE_DIR_PATH = "excel/output"; //-> /home/zhyun/my-projects/study-excel-and-zip-file-control/excel/output


    /**
     * 엑셀 생성
     */
    public void create(String filename, boolean isXlsx) {
        String nowDateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyMMddHHmmss"));
        String fileExtension = isXlsx ? "xlsx" : "xls";
        String newFilename = "%s/%s_%s.%s".formatted(NEW_FILE_DIR_PATH, filename, nowDateTime, fileExtension);

        try (Workbook wb = WorkbookFactory.create(isXlsx)) {
            Sheet sheet = wb.createSheet("new sheet");

            // 임시 데이터 생성
            List<List<String>> rowData = getData();

            // 엑셀 작성
            for (int i = 0; i < rowData.size(); i++) {
                Row row = sheet.createRow(i);
                List<String> cols = rowData.get(i);
                for (int j = 0; j < cols.size(); j++) {
                    try {
                        row.createCell(j).setCellValue(
                                Integer.parseInt(cols.get(j))
                        );
                    } catch (Exception e) {
                        try {
                            row.createCell(j).setCellValue(
                                    Double.parseDouble(cols.get(j))
                            );
                        } catch (Exception ee) {
                            row.createCell(j).setCellValue(
                                    cols.get(j)
                            );
                        }
                    }
                }
            }

            // 파일 생성
            makeDir();
            try  (OutputStream fileOut = new FileOutputStream(newFilename)) {
                wb.write(fileOut);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 암호화 엑셀 읽기
     */
    public void printAllSecret(String filename, String password) {
        printFilename(filename);

        try (FileInputStream fis = new FileInputStream(FILE_DIR_PATH.formatted(filename));
             Workbook workbook = WorkbookFactory.create(fis, password)) {

            readWorkbook(workbook);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 전체 내용 출력
     * - <a href="https://poi.apache.org/components/spreadsheet/quick-guide.html#ReadWriteWorkbook">poi.apache#ReadWriteWorkbook</a>
     */
    public void printAll(String filename) {
        printFilename(filename);

        try (FileInputStream fis = new FileInputStream(FILE_DIR_PATH.formatted(filename));
             Workbook wb = WorkbookFactory.create(fis)) {

            readWorkbook(wb);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }





    private void makeDir() {
        Path path = Paths.get(NEW_FILE_DIR_PATH);
        try {
            Files.createDirectories(path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void readWorkbook(Workbook wb) {
        DataFormatter formatter = new DataFormatter();

        for (Sheet sheet : wb) {
            System.out.printf("⭐\uFE0F sheet name: %s ==============================\n", sheet.getSheetName());

            for (Row row : sheet) {
                System.out.printf("%5d row ==============================\n", row.getRowNum());

                short firstCellNum = row.getFirstCellNum();
                short lastCellNum = row.getLastCellNum();
                for (int i = firstCellNum; i < lastCellNum; i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String value = formatter.formatCellValue(cell);
                    System.out.printf("%s\n", value);
                }

                System.out.println();
            }
        }
    }

    private void printFilename(String filename) {
        System.out.printf("\n\n\uD83D\uDCBE filename: %s ==================%n", filename);
    }

    /**
     * 임시 데이터 생성
     */
    private List<List<String>> getData() {
        List<List<String>> rowData = new ArrayList<>();

        rowData.add(List.of("header1", "header2", "header3", "header4", "header5", "header6", "header7"));

        List<String> colData = new ArrayList<>();
        colData.add("1");
        colData.add("이");
        colData.add("sam");
        colData.add("4.0");
        colData.add("5678910");
        colData.add("11.1213141516171819");

        for (int i = 0; i < 100; i++) {
            rowData.add(colData);
        }

        return rowData;
    }
}
