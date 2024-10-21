package excel;

public class Main {
    static final String FILENAME_XLSX = "test.xlsx";
    static final String FILENAME_XLS = "test.xls";


    public static void main(String[] args) {
        ExcelService excel = new ExcelService();

        excel.printAll(FILENAME_XLSX);
        excel.printAll(FILENAME_XLS);






    }
}
