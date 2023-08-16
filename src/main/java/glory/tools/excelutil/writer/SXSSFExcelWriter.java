package glory.tools.excelutil.writer;


import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

import glory.tools.excelutil.exception.ExcelInternalException;
import glory.tools.excelutil.resource.DataFormatDecider;
import glory.tools.excelutil.resource.ExcelRenderLocation;
import glory.tools.excelutil.resource.ExcelRenderResource;
import glory.tools.excelutil.resource.ExcelRenderResourceFactory;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public abstract class SXSSFExcelWriter<T> implements ExcelWriter<T> {

    protected static final int                HEADER_ROW_INDEX   = 0;
    protected static final int                COLUMN_START_INDEX = 0;
    protected static final SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;

    protected SXSSFWorkbook       wb;
    protected Sheet               sheet;
    protected ExcelRenderResource resource;

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

    protected abstract void validateData(List<T> data);

    public abstract void renderExcel(List<T> data);

    protected void renderHeaders() {
        Row row = sheet.createRow(HEADER_ROW_INDEX);

        int columnIndex = COLUMN_START_INDEX;
        for (String dataFieldName : resource.getDataFieldNames()) {
            sheet.setColumnWidth(columnIndex, resource.getExcelHeaderWidth(dataFieldName) * 256);

            Cell cell = row.createCell(columnIndex++);
            cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.HEADER));
            cell.setCellValue(resource.getExcelHeaderName(dataFieldName));
        }
    }

    protected void renderDataRow(T data, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = 0;
        for (String dataFieldName : resource.getDataFieldNames()) {
            Cell cell = row.createCell(columnIndex++);

            try {
                Field field = getField(data, dataFieldName);

                if (StringUtils.equals(resource.getCountColumnName(), field.getName())) {
                    setCellValue(cell, rowIndex);
                } else {
                    Object cellValue = field.get(data);
                    setCellValue(cell, cellValue);
                }

            } catch (IllegalAccessException e) {
                throw new ExcelInternalException(e.getMessage(), e);
            }
        }
    }

    private Field getField(T data, String fieldName) {
        try {
            return FieldUtils.getField(data.getClass(), fieldName, true);

        } catch (Exception e) {
            throw new ExcelInternalException(e.getMessage(), e);
        }
    }

    private void setCellValue(Cell cell, Object value) {
        if (value instanceof Number castedValue) {
            cell.setCellValue(castedValue.doubleValue());
        } else if (value instanceof Boolean castedValue) {
            cell.setCellValue(castedValue);
        } else if (value instanceof LocalDate castedValue) {
            cell.setCellValue(castedValue);
        } else if (value instanceof LocalDateTime castedValue) {
            cell.setCellValue(castedValue);
        } else {
            cell.setCellValue(value == null ? "" : value.toString());
        }
    }

}
