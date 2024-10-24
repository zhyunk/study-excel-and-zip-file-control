package kim.zhyun.study.excel;

import org.apache.poi.ss.usermodel.*;
import kim.zhyun.study.util.FileUtil;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import static kim.zhyun.study.Main.FILE_DIR_PATH;


public class ExcelService {
    private final String NEW_FILE_DIR_PATH = "sampleFile/output"; //-> /home/zhyun/my-projects/study-excel-and-zip-file-control/sampleFile/output

    private final DataFormatter formatter = new DataFormatter();


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

            // 폴더 생성
            FileUtil.makeDir(NEW_FILE_DIR_PATH);

            // 파일 생성
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
        FileUtil.printFilename(filename);

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
        FileUtil.printFilename(filename);

        try (FileInputStream fis = new FileInputStream(FILE_DIR_PATH.formatted(filename));
             Workbook wb = WorkbookFactory.create(fis)) {

            readWorkbook(wb);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }





    private void readWorkbook(Workbook wb) {
        for (Sheet sheet : wb) {
            System.out.printf("⭐\uFE0F sheet name: %s ==============================\n", sheet.getSheetName());

            for (Row row : sheet) {
                System.out.printf("%5d row ==============================\n", row.getRowNum());

                Row rowHeader = sheet.getRow(0);
                short firstCellNum = row.getFirstCellNum();
                short lastCellNum = row.getLastCellNum();
                for (int i = firstCellNum; i < lastCellNum; i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String value = getValue(cell);

                    Cell cellHeader = rowHeader.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String valueHeader = formatter.formatCellValue(cellHeader);

                    System.out.printf("%s: %s\n", valueHeader, value);
                }

                System.out.println();
            }
        }
    }

    /**
     * 숫자와 날짜 구분 후 문자열로 반환
     */
    private String getValue(Cell cell) {
        boolean isNumber = cell != null ? cell.getCellType() == CellType.NUMERIC : false;
        boolean isDateFormatted = isNumber ? DateUtil.isCellDateFormatted(cell) : false;
        boolean isNotDate = isNumber && !isDateFormatted;

        return isNotDate ?
                cell.getNumericCellValue() == Math.floor(cell.getNumericCellValue()) ?
                        BigDecimal.valueOf((long) cell.getNumericCellValue()).toPlainString() :
                        BigDecimal.valueOf(cell.getNumericCellValue()).toPlainString() :
                formatter.formatCellValue(cell);
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
