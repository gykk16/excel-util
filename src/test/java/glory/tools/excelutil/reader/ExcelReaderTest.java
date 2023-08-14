package glory.tools.excelutil.reader;

import static org.assertj.core.api.Assertions.assertThat;

import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class ExcelReaderTest {

    @DisplayName("Read excel file")
    @Test
    void read_excel() throws Exception {
        // given
        String fileLocation = "src/test/resources/static/test/temp/excelfile/test.xlsx";
        boolean hasHeader = true;

        // when
        Map<Integer, List<String>> data = ExcelReader.readExcel(fileLocation, hasHeader);

        // then
        assertThat(data).hasSize(2)
                .extractingByKeys(1, 2)
                .containsExactlyInAnyOrder(
                        List.of("1", "test_0", "0"),
                        List.of("2", "test_1", "1")
                );
    }
}