package kim.zhyun.study.zip;

import kim.zhyun.study.excel.CsvService;
import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.exception.ZipException;

import java.io.IOException;

import static kim.zhyun.study.Main.FILE_DIR_PATH;

public class ZipService {

    public void printAll(String filename, String password) {
        String readFilePath = FILE_DIR_PATH.formatted(filename);
        String writeFilePath = "sampleFile/output"; //-> /home/zhyun/my-projects/study-excel-and-zip-file-control/sampleFile/output

        try (ZipFile zipFile = new ZipFile(readFilePath)) {

            if (zipFile.isEncrypted() && password != null) {
                zipFile.setPassword(password.toCharArray());
            }

            zipFile.extractAll(writeFilePath);

            CsvService csvService = new CsvService();
            csvService.printAll("output/qweqwe_20241022_23_zxczxc.csv");

        } catch (ZipException e) {
            System.err.println("Error extracting files: " + e.getMessage());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }
}
