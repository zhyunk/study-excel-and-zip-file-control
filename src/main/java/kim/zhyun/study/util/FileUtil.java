package kim.zhyun.study.util;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class FileUtil {

    /**
     * 파일 저장할 폴더 없으면 생성
     */
    public static void makeDir(String NEW_FILE_DIR_PATH) {
        Path path = Paths.get(NEW_FILE_DIR_PATH);
        try {
            Files.createDirectories(path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 파일 제목 콘솔에 출력
     */
    public static void printFilename(String filename) {
        System.out.printf("\n\n\uD83D\uDCBE filename: %s ==================%n", filename);
    }

}
