package kim.zhyun.study.excel;

import kim.zhyun.study.util.FileUtil;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

import static kim.zhyun.study.Main.FILE_DIR_PATH;


public class CsvService {

    /**
     * csv 전체 내용 출력
     */
    public void printAll(String filename) {
        FileUtil.printFilename(filename);

        String line;
        String csvSplitBy = ",";

        ArrayList<ArrayList<String>> rows = new ArrayList<>();

        try (BufferedReader br = new BufferedReader(new FileReader(FILE_DIR_PATH.formatted(filename)))) {
            while ((line = br.readLine()) != null) {
                String[] values = line.split(csvSplitBy);
                rows.add(new ArrayList<>(Arrays.asList(values)));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        // print
        int rowsCount = rows.size();
        for (int j = 1; j < rowsCount; j++) {
            for (int i = 0; i < rows.get(0).size(); i++) {
                System.out.printf("%s: %s\n", rows.get(0).get(i), rows.get(j).get(i));
            }
            System.out.println();
        }
    }
}
