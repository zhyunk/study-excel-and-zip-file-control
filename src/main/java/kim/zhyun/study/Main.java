package kim.zhyun.study;

import kim.zhyun.study.excel.CsvService;
import kim.zhyun.study.excel.ExcelService;
import kim.zhyun.study.zip.ZipService;

public class Main {
    public static final String FILE_DIR_PATH = "sampleFile/%s";

    static final String FILENAME_XLSX = "test.xlsx";
    static final String FILENAME_XLS = "test.xls";
    static final String FILENAME_PASSWORD_XLSX = "test-password.xlsx";
    static final String FILENAME_PASSWORD_XLS = "test-password.xls";
    static final String FILENAME_PASSWORD_XLSX_2 = "test-password-from-prod.xlsx";
    static final String FILENAME_PASSWORD_XLS_2 = "test-password-from-prod.xls";
    static final String FILENAME_CSV = "test.csv";
    static final String FILENAME_ZIP = "asdasd_123123_123123.zip";


    public static void main(String[] args) {
        ExcelService excel = new ExcelService();

//        excel.printAll(FILENAME_XLSX);
//        excel.printAll(FILENAME_XLS);

//        excel.printAllSecret(FILENAME_PASSWORD_XLS, "1234");
//        excel.printAllSecret(FILENAME_PASSWORD_XLSX, "1234");
//        excel.printAllSecret(FILENAME_PASSWORD_XLS_2, "qweqwe123123");
//        excel.printAllSecret(FILENAME_PASSWORD_XLSX_2, "qweqwe123123");

//        excel.create("test", true);

        CsvService csv = new CsvService();
//        csv.printAll(FILENAME_CSV);

        ZipService zip = new ZipService();
        zip.printAll(FILENAME_ZIP, "qweqwe123123");



    }
}
