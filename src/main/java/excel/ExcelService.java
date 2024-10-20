package excel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;

public class ExcelService {
    private final String FILE_DIR_PATH = "/home/zhyun/my-projects/study-excel-and-zip-file-control/excel/%s";


    /**
     * 전체 내용 출력
     * - <a href="https://poi.apache.org/components/spreadsheet/quick-guide.html#ReadWriteWorkbook">poi.apache#ReadWriteWorkbook</a>
     */
    public void printAll(String filename) {
        printFilename(filename);

        try (InputStream inp = new FileInputStream(FILE_DIR_PATH.formatted(filename))) {
            Workbook wb = WorkbookFactory.create(inp);

            for (Sheet sheet : wb) {
                System.out.printf("⭐\uFE0F sheet name: %s ==============================\n", sheet.getSheetName());

                for (Row row : sheet) {
                    System.out.printf("%5d row ==============================\n", row.getRowNum());

                    for (Cell cell : row) {
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case STRING  -> System.out.println(cell.getStringCellValue());
                            case NUMERIC -> System.out.println(BigDecimal.valueOf(cell.getNumericCellValue()).toPlainString());
                            case BOOLEAN -> System.out.println(cell.getBooleanCellValue());
                            case ERROR   -> System.out.printf("error: %s\n", cell.getErrorCellValue());
                        }
                    }

                    System.out.println();
                }
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void printFilename(String filename) {
        System.out.printf("\n\n\uD83D\uDCBE filename: %s ==================%n", filename);
    }






    /**
     * format
     * -
     */
    public void method(String filename) {
        printFilename(filename);

        try (InputStream inp = new FileInputStream(FILE_DIR_PATH.formatted(filename))) {
            Workbook wb = WorkbookFactory.create(inp);



        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
