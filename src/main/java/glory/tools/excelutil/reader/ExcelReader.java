package glory.tools.excelutil.reader;


import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.ToIntFunction;

import glory.tools.excelutil.exception.ExcelException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

    private static final int LIMIT_ROW = SpreadsheetVersion.EXCEL2007.getMaxRows() - 1_000;

    private ExcelReader() {
    }

    public static Map<Integer, List<String>> readExcel(String fileLocation, boolean hasHeader) {

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(fileLocation))) {
            Sheet sheet = workbook.getSheetAt(0);
            validateRowLimit(sheet);

            var lastRowNum = sheet.getLastRowNum();
            var lastCellNum = sheet.getRow(0).getLastCellNum();
            var startRow = hasHeader ? 1 : 0;

            Map<Integer, List<String>> data = new HashMap<>();

            for (int currentRowNum = startRow; currentRowNum < lastRowNum + 1; currentRowNum++) {
                data.put(currentRowNum, new ArrayList<>());
                Row row = sheet.getRow(currentRowNum);
                rowToMap(data, lastCellNum, currentRowNum, row);
            }

            return data;

        } catch (IOException e) {
            throw new ExcelException("Excel processing error.", e);

        }

    }

    public static int readExcelAndApplyFunction(String fileLocation, boolean hasHeader, int functionBatchSize,
            ToIntFunction<Map<Integer, List<String>>> function) {

        int processed = 0;
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(fileLocation))) {
            Sheet sheet = workbook.getSheetAt(0);
            validateRowLimit(sheet);

            var lastRowNum = sheet.getLastRowNum();
            var lastCellNum = sheet.getRow(0).getLastCellNum();
            var startRow = hasHeader ? 1 : 0;

            Map<Integer, List<String>> data = new HashMap<>();

            for (int currentRowNum = startRow; currentRowNum < lastRowNum + 1; currentRowNum++) {

                if (currentRowNum != 0 && currentRowNum % functionBatchSize == 0) {
                    /* Apply function */
                    processed += function.applyAsInt(data);
                    data.clear();
                }

                data.put(currentRowNum, new ArrayList<>());
                Row row = sheet.getRow(currentRowNum);
                rowToMap(data, lastCellNum, currentRowNum, row);
            }

            /* Apply function */
            processed += function.applyAsInt(data);
            data.clear();

            return processed;

        } catch (IOException e) {
            throw new ExcelException("Excel processing error.", e);

        }
    }

    private static void validateRowLimit(Sheet sheet) {
        if (sheet.getLastRowNum() > LIMIT_ROW) {
            throw new ExcelException("Excel max row is " + LIMIT_ROW + " row.");
        }
    }

    private static void rowToMap(Map<Integer, List<String>> data, short lastCellNum, int currentRowNum, Row row) {
        for (int currCellNum = 0; currCellNum < lastCellNum; currCellNum++) {
            Cell cell = row.getCell(currCellNum);
            String value = extractCellValue(cell);
            data.get(currentRowNum).add(value);
        }
    }

    private static String extractCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        return switch (cell.getCellType()) {
            case STRING -> cell.getRichStringCellValue().getString().trim();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? cell.getLocalDateTimeCellValue().toString()
                    : String.valueOf((long)cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }
}
