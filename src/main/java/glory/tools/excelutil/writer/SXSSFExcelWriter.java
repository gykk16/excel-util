package glory.tools.excelutil.writer;


import static org.apache.commons.lang3.reflect.FieldUtils.getField;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Collections;
import java.util.List;


import glory.tools.excelutil.exception.ExcelInternalException;
import glory.tools.excelutil.resource.DataFormatDecider;
import glory.tools.excelutil.resource.DefaultDataFormatDecider;
import glory.tools.excelutil.resource.ExcelRenderLocation;
import glory.tools.excelutil.resource.ExcelRenderResource;
import glory.tools.excelutil.resource.ExcelRenderResourceFactory;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public abstract class SXSSFExcelWriter<T> implements ExcelWriter<T> {

    protected static final SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;

    protected SXSSFWorkbook       wb;
    protected Sheet               sheet;
    protected ExcelRenderResource resource;


    /**
     * SXSSFExcelFile
     *
     * @param type Class type to be rendered
     */
    protected SXSSFExcelWriter(Class<T> type) {
        this(Collections.emptyList(), type, new DefaultDataFormatDecider());
    }

    /**
     * SXSSFExcelFile
     *
     * @param data List Data to render excel file. data should have at least one @ExcelColumn on fields
     * @param type Class type to be rendered
     */
    protected SXSSFExcelWriter(List<T> data, Class<T> type) {
        this(data, type, new DefaultDataFormatDecider());
    }

    /**
     * SXSSFExcelFile
     *
     * @param data              List Data to render excel file. data should have at least one @ExcelColumn on fields
     * @param type              Class type to be rendered
     * @param dataFormatDecider Custom DataFormatDecider
     */
    protected SXSSFExcelWriter(List<T> data, Class<T> type, DataFormatDecider dataFormatDecider) {
        this.wb = new SXSSFWorkbook();
        this.resource = ExcelRenderResourceFactory.prepareRenderResource(type, wb, dataFormatDecider);
        // renderExcel(data);
    }

    protected abstract void validateData(List<T> data);

    protected abstract void renderExcel(List<T> data);

    protected void renderHeadersWithNewSheet(Sheet sheet, int rowIndex, int columnStartIndex) {
        Row row = sheet.createRow(rowIndex);

        int columnIndex = columnStartIndex;
        for (String dataFieldName : resource.getDataFieldNames()) {

            sheet.setColumnWidth(columnIndex, resource.getExcelHeaderWidth(dataFieldName) * 256);

            Cell cell = row.createCell(columnIndex++);
            cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.HEADER));
            cell.setCellValue(resource.getExcelHeaderName(dataFieldName));
        }
    }

    protected void renderBody(Object data, int rowIndex, int columnStartIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = columnStartIndex;
        for (String dataFieldName : resource.getDataFieldNames()) {
            Cell cell = row.createCell(columnIndex++);
            try {
                Field field = getField(data.getClass(), dataFieldName, true);
                // field.setAccessible(true);
                cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.BODY));

                if (StringUtils.equals(resource.getCountColumnName(), field.getName())) {
                    renderCellValue(cell, rowIndex);
                } else {
                    Object cellValue = field.get(data);
                    renderCellValue(cell, cellValue);
                }

            } catch (Exception e) {
                throw new ExcelInternalException(e.getMessage(), e);
            }
        }
    }

    private void renderCellValue(Cell cell, Object cellValue) {
        // number
        if (cellValue instanceof Number numberValue) {
            cell.setCellValue(numberValue.doubleValue());
            return;
        }
        // bool
        if (cellValue instanceof Boolean booleanValue) {
            cell.setCellValue(booleanValue);
            return;
        }
        // localDate
        if (cellValue instanceof LocalDate localDateValue) {
            cell.setCellValue(localDateValue);
            return;
        }
        // localDateTime
        if (cellValue instanceof LocalDateTime localDateTimeValue) {
            cell.setCellValue(localDateTimeValue);
            return;
        }
        // 나머지
        cell.setCellValue(cellValue == null ? "" : cellValue.toString());
    }

    public void write(OutputStream stream) throws IOException {
        try (stream) {
            wb.write(stream);
            wb.close();
            wb.dispose();
        }
    }

    public SpreadsheetVersion getSupplyExcelVersion() {
        return supplyExcelVersion;
    }

}
