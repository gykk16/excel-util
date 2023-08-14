package glory.tools.excelutil.writer.singlesheet;


import java.util.Collections;
import java.util.List;

import glory.tools.excelutil.resource.DataFormatDecider;
import glory.tools.excelutil.resource.DefaultDataFormatDecider;
import glory.tools.excelutil.writer.SXSSFExcelWriter;
import org.springframework.lang.NonNull;


/**
 * SingleSheetExcelWriter
 * <p>
 * <li>support Excel Version over 2007</li>
 * <li>supports only one sheet</li>
 * <li>support Different DataFormat by Class Type</li>
 * <li>support Custom CellStyle according to (header or body) and data field</li>
 * </p>
 */
public class SingleSheetExcelWriter<T> extends SXSSFExcelWriter<T> {

    private static final int ROW_START_INDEX    = 0;
    private static final int COLUMN_START_INDEX = 0;
    private              int currentRowIndex    = ROW_START_INDEX;

    public SingleSheetExcelWriter(@NonNull Class<T> type) {
        this(Collections.emptyList(), type);
    }

    public SingleSheetExcelWriter(@NonNull List<T> data, @NonNull Class<T> type) {
        this(data, type, new DefaultDataFormatDecider());
    }

    public SingleSheetExcelWriter(@NonNull List<T> data, @NonNull Class<T> type,
            @NonNull DataFormatDecider dataFormatDecider) {
        super(data, type, dataFormatDecider);
        this.validateData(data);
        this.renderExcel(data);
    }

    @Override
    protected void validateData(List<T> data) {
        int maxRows = supplyExcelVersion.getMaxRows();
        if (data.size() >= maxRows) {
            throw new IllegalArgumentException(
                    "This concrete ExcelFile does not support over %s rows".formatted(maxRows));
        }
    }

    @Override
    public void renderExcel(List<T> data) {
        // 1. Create sheet and renderHeader
        sheet = wb.createSheet();
        renderHeadersWithNewSheet(sheet, currentRowIndex++, COLUMN_START_INDEX);

        if (data.isEmpty()) {
            return;
        }

        // 2. Render Body
        for (Object renderedData : data) {
            renderBody(renderedData, currentRowIndex++, COLUMN_START_INDEX);
        }
    }

    @Override
    public void addRows(List<T> data) {
        for (Object renderedData : data) {
            renderBody(renderedData, currentRowIndex++, COLUMN_START_INDEX);
        }
    }

}
