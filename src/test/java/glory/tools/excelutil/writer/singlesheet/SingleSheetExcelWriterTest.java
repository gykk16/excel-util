package glory.tools.excelutil.writer.singlesheet;

import static glory.tools.excelutil.common.ExcelFileUtil.*;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

import glory.tools.excelutil.common.ExcelFileUtil;
import glory.tools.excelutil.reader.ExcelReader;
import glory.tools.excelutil.writer.ExcelWriter;
import glory.tools.excelutil.writer.TestDataDTO;
import org.apache.poi.ss.SpreadsheetVersion;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class SingleSheetExcelWriterTest {

    @DisplayName("Initialization of SingleSheetExcelWriter")
    @Test
    void init() throws Exception {
        // when
        ExcelWriter<TestDataDTO> writer = new SingleSheetExcelWriter<>(TestDataDTO.class);

        // then
        assertThat(writer).isNotNull();
    }

    @Disabled("This test is disabled because it takes too long to run")
    @DisplayName("SingleSheetExcel validate max data size")
    @Test
    void validate_data() throws Exception {
        // given
        SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;
        List<TestDataDTO> dataDTOS = IntStream.range(0, supplyExcelVersion.getMaxRows() - 1)
                .mapToObj(i -> new TestDataDTO("test_" + i, i))
                .toList();

        // when
        ExcelWriter<TestDataDTO> writer = new SingleSheetExcelWriter<>(dataDTOS, TestDataDTO.class);

        // then
        assertThat(writer).isNotNull();
    }

    @Disabled("This test is disabled because it takes too long to run")
    @DisplayName("SingleSheetExcel throw exception if max data size is exceeded")
    @Test
    void validate_data_exception() throws Exception {
        // given
        SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;
        List<TestDataDTO> dataDTOS = IntStream.range(0, supplyExcelVersion.getMaxRows())
                .mapToObj(i -> new TestDataDTO("test_" + i, i))
                .toList();

        // when
        assertThatThrownBy(() -> new SingleSheetExcelWriter<>(dataDTOS, TestDataDTO.class))
                // then
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessage(
                        "This concrete ExcelFile does not support over " + supplyExcelVersion.getMaxRows() + " rows");
    }

    @DisplayName("Render excel file")
    @Test
    void render_excel() throws Exception {
        // given
        final String savePath = "src/test/resources/static/test/temp";
        final String filename = "test";

        ExcelWriter<TestDataDTO> writer = new SingleSheetExcelWriter<>(TestDataDTO.class);

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

}