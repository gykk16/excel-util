package glory.tools.excelutil.writer.multisheet;

import static glory.tools.excelutil.common.ExcelFileUtil.EXCEL_FILE_EXTENSION;
import static glory.tools.excelutil.common.ExcelFileUtil.getExcelFileWriter;
import static org.assertj.core.api.Assertions.assertThat;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

import glory.tools.excelutil.reader.ExcelReader;
import glory.tools.excelutil.writer.ExcelWriter;
import glory.tools.excelutil.writer.TestDataDTO;
import org.apache.poi.ss.SpreadsheetVersion;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class MultiSheetExcelWriterTest {

    @DisplayName("Initialization of MultiSheetExcelWriter")
    @Test
    void init() throws Exception {
        // when
        ExcelWriter<TestDataDTO> writer = new MultiSheetExcelWriter<>(TestDataDTO.class);

        // then
        assertThat(writer).isNotNull();
    }

    @DisplayName("Render excel file")
    @Test
    void render_excel() throws Exception {
        // given
        final String savePath = "src/test/resources/static/test/temp";
        final String filename = "test";

        ExcelWriter<TestDataDTO> writer = new MultiSheetExcelWriter<>(TestDataDTO.class);

        List<TestDataDTO> dataDTOS = IntStream.range(0, 2)
                .mapToObj(i -> new TestDataDTO("test_" + i, i))
                .toList();

        // when
        writer.addRows(dataDTOS);
        writer.write(getExcelFileWriter(savePath, filename));

        // then
        File saved = new File(savePath, filename + EXCEL_FILE_EXTENSION);
        assertThat(saved).exists();

        Map<Integer, List<String>> readExcel = ExcelReader.readExcel(saved.toString(), true);
        assertThat(readExcel).hasSize(2)
                .extractingByKeys(1, 2)
                .containsExactlyInAnyOrder(
                        List.of("1", "test_0", "0"),
                        List.of("2", "test_1", "1")
                );

        saved.deleteOnExit();
    }

    @Disabled("This test is disabled because it takes too long to run")
    @DisplayName("Render multi sheet excel file")
    @Test
    void render_multi_sheet_excel() throws Exception {
        // given
        final String savePath = "src/test/resources/static/test/temp";
        final String filename = "test";

        SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;

        ExcelWriter<TestDataDTO> writer = new MultiSheetExcelWriter<>(TestDataDTO.class);

        List<TestDataDTO> dataDTOS = IntStream.range(0, supplyExcelVersion.getMaxRows() + 2)
                .mapToObj(i -> new TestDataDTO("test_" + i, i))
                .toList();

        // when
        writer.addRows(dataDTOS);
        writer.write(getExcelFileWriter(savePath, filename));

        // then
        File saved = new File(savePath, filename + EXCEL_FILE_EXTENSION);
        assertThat(saved).exists();

        saved.deleteOnExit();
    }
}