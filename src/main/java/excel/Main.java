package excel;

public class Main {

    public static void main(String[] args) {
        String filename1 = "test.xlsx";
        String filename2 = "test.xls";

        ExcelService excel = new ExcelService();
        excel.printAll(filename1);
        excel.printAll(filename2);
    }
}
