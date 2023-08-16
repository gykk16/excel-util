package glory.tools.excelutil.writer.multisheet;


import java.util.Collections;
import java.util.List;

import glory.tools.excelutil.resource.DataFormatDecider;
import glory.tools.excelutil.resource.DefaultDataFormatDecider;
import glory.tools.excelutil.writer.SXSSFExcelWriter;
import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.springframework.lang.NonNull;

/**
 * MultiSheetExcelWriter
 * <p>
 * <li>support Excel Version over 2007</li>
 * <li>support multi sheet rendering</li>
 * <li>support Different DataFormat by Class Type</li>
 * <li>support Custom CellStyle according to (header or body) and data field</li>
 * </p>
 */
public class MultiSheetExcelWriter<T> extends SXSSFExcelWriter<T> {

    private static final int MAX_ROW_CAN_BE_RENDERED = supplyExcelVersion.getMaxRows() - 48_576;
    private static final int ROW_START_INDEX         = 1; // 0 is for header
    private              int currentRowIndex         = ROW_START_INDEX;

    public MultiSheetExcelWriter(@NonNull Class<T> type) {
        this(Collections.emptyList(), type);
    }

    public MultiSheetExcelWriter(@NonNull List<T> data, @NonNull Class<T> type) {
        this(data, type, new DefaultDataFormatDecider());
    }

    public MultiSheetExcelWriter(@NonNull List<T> data, @NonNull Class<T> type,
            @NonNull DataFormatDecider dataFormatDecider) {
        super(data, type, dataFormatDecider);
        this.setZipMode();
        this.renderExcel(data);
    }

    @Override
    public void renderExcel(List<T> data) {

        createNewSheetWithHeader();

        if (data.isEmpty()) {
            return;
        }

        addRows(data);
    }

    @Override
    public void addRows(List<T> data) {
        for (T rowData : data) {
            if (currentRowIndex == MAX_ROW_CAN_BE_RENDERED) {
                currentRowIndex = 1;
                createNewSheetWithHeader();
            }
            renderDataRow(rowData, currentRowIndex++);
        }
    }

    @Override
    protected void validateData(List<T> data) {
        // do nothing
    }

    private void setZipMode() {
        wb.setZip64Mode(Zip64Mode.Always);
    }

    private void createNewSheetWithHeader() {
        sheet = wb.createSheet();
        renderHeaders();
    }

}
