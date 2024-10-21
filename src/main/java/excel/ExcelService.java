package excel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelService {
    private final String FILE_DIR_PATH = "/home/zhyun/my-projects/study-excel-and-zip-file-control/excel/%s";


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
}
