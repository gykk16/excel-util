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

    private static final int maxRowCanBeRendered = supplyExcelVersion.getMaxRows() - 48_576;
    private static final int ROW_START_INDEX     = 0;
    private static final int COLUMN_START_INDEX  = 0;
    private              int currentRowIndex     = ROW_START_INDEX;

    public MultiSheetExcelWriter(@NonNull Class<T> type) {
        this(Collections.emptyList(), type);
        wb.setZip64Mode(Zip64Mode.Always);
    }

    /*
     * If you use SXSSF with huge data, you need to set zip mode
     * see http://apache-poi.1045710.n5.nabble.com/Bug-62872-New-Writing-large-files-with-800k-rows-gives-java-io-IOException-This-archive-contains-unc-td5732006.html
     */
    public MultiSheetExcelWriter(@NonNull List<T> data, @NonNull Class<T> type) {
        this(data, type, new DefaultDataFormatDecider());
        wb.setZip64Mode(Zip64Mode.Always);
    }

    public MultiSheetExcelWriter(@NonNull List<T> data, @NonNull Class<T> type,
            @NonNull DataFormatDecider dataFormatDecider) {
        super(data, type, dataFormatDecider);
        wb.setZip64Mode(Zip64Mode.Always);
        this.renderExcel(data);
    }

    @Override
    protected void validateData(List<T> data) {
        // do nothing
    }

    @Override
    protected void renderExcel(List<T> data) {
        // 1. Create header and return if data is empty
        if (data.isEmpty()) {
            createNewSheetWithHeader();
            return;
        }

        // 2. Render body
        createNewSheetWithHeader();
        addRows(data);
    }

    @Override
    public void addRows(List<T> data) {
        for (Object renderedData : data) {
            renderBody(renderedData, currentRowIndex++, COLUMN_START_INDEX);
            if (currentRowIndex == maxRowCanBeRendered) {
                currentRowIndex = 0;
                createNewSheetWithHeader();
            }
        }
    }

    private void createNewSheetWithHeader() {
        sheet = wb.createSheet();
        renderHeadersWithNewSheet(sheet, ROW_START_INDEX, COLUMN_START_INDEX);
        currentRowIndex++;
    }

}
