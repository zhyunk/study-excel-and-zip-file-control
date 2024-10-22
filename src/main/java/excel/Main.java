package excel;

public class Main {
    static final String FILENAME_XLSX = "test.xlsx";
    static final String FILENAME_XLS = "test.xls";
    static final String FILENAME_PASSWORD_XLSX = "test-password.xlsx";
    static final String FILENAME_PASSWORD_XLS = "test-password.xls";
    static final String FILENAME_PASSWORD_XLSX_2 = "test-password-from-prod.xlsx";
    static final String FILENAME_PASSWORD_XLS_2 = "test-password-from-prod.xls";


    public static void main(String[] args) {
        ExcelService excel = new ExcelService();

//        excel.printAll(FILENAME_XLSX);
//        excel.printAll(FILENAME_XLS);

//        excel.printAllSecret(FILENAME_PASSWORD_XLS, "1234");
//        excel.printAllSecret(FILENAME_PASSWORD_XLSX, "1234");
        excel.printAllSecret(FILENAME_PASSWORD_XLS_2, "qweqwe123123");
        excel.printAllSecret(FILENAME_PASSWORD_XLSX_2, "qweqwe123123");

//        excel.create("test", true);






    }
}
