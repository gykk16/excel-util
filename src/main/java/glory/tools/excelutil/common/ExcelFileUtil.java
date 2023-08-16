package glory.tools.excelutil.common;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.springframework.lang.NonNull;
import org.springframework.util.Assert;

public class ExcelFileUtil {

    public static final String EXCEL_FILE_EXTENSION = ".xlsx";

    private ExcelFileUtil() {
    }

    /**
     * Returns true if the file extension is xlsx
     *
     * @param fileName file name with extension
     * @return true if the file extension is xlsx
     */
    public static boolean isExcelFile(String fileName) {
        return fileName.toLowerCase().endsWith(EXCEL_FILE_EXTENSION);
    }

    /**
     * Returns OutputStream of excel file
     *
     * @param savePath directory path to save excel file
     * @param filename file name without extension
     * @return OutputStream of excel file
     */
    public static OutputStream getExcelFileWriter(@NonNull String savePath, String filename) {
        Assert.notNull(savePath, "savePath must not be null");
        createDirIfNotExists(savePath);

        String fileName = filename + EXCEL_FILE_EXTENSION;

        File outputFile;
        try {
            outputFile = new File(savePath, fileName);
            return new FileOutputStream(outputFile);

        } catch (FileNotFoundException e) {
            throw new IllegalStateException("File not found.", e);
        }
    }

    private static void createDirIfNotExists(String dirPath) {
        File dir = new File(dirPath);
        if (!dir.exists()) {
            dir.mkdirs();
        }
    }

}
