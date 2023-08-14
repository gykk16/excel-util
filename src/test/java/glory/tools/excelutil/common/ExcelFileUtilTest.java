package glory.tools.excelutil.common;

import static glory.tools.excelutil.common.ExcelFileUtil.EXCEL_FILE_EXTENSION;
import static glory.tools.excelutil.common.ExcelFileUtil.getExcelFileWriter;
import static glory.tools.excelutil.common.ExcelFileUtil.isExcelFile;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class ExcelFileUtilTest {

    private static final String TEST_RESOURCES_LOC = "src/test/resources/static/test/temp";

    @DisplayName("Returns true if the file extension is xlsx")
    @Test
    void is_excel_file() throws Exception {
        assertThat(isExcelFile("test.xlsx")).isTrue();
        assertThat(isExcelFile("test.XLSX")).isTrue();
        assertThat(isExcelFile("test.xlx")).isFalse();
    }

    @DisplayName("Returns OutputStream of excel file")
    @Test
    void get_excel_file_writer() throws Exception {
        // given
        String filename = "test";
        Path tempPath = Files.createTempDirectory("excelTest");

        // when
        OutputStream os = getExcelFileWriter(tempPath.toString(), filename);

        // then
        assertThat(os).isNotNull();
        assertThat(Files.exists(tempPath.resolve(filename + EXCEL_FILE_EXTENSION))).isTrue();
    }

    @DisplayName("Throws Exception if the savePath is null")
    @Test
    void get_excel_file_writer_null() throws Exception {
        // given

        // when
        assertThatThrownBy(() -> getExcelFileWriter(null, "test"))
                // then
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("savePath must not be null");
    }
}