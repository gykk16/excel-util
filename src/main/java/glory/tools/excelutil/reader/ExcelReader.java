package glory.tools.excelutil.reader;


import static glory.tools.excelutil.common.LogStringUtil.logSubTitle;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.ToIntFunction;

import glory.tools.excelutil.exception.ExcelException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StopWatch;

public class ExcelReader {

    private static final Logger log = LoggerFactory.getLogger(ExcelReader.class);

    private static final int LIMIT_ROW = 100_100;

    private ExcelReader() {
    }

    public static Map<Integer, List<String>> readExcel(String fileLocation, boolean hasHeader) {
        StopWatch stopWatch = new StopWatch();
        Map<Integer, List<String>> map = readExcel(fileLocation, hasHeader, stopWatch);
        log.debug("==> {}", stopWatch);
        return map;
    }

    public static Map<Integer, List<String>> readExcel(String fileLocation, boolean hasHeader, StopWatch stopWatch) {
        log.debug(logSubTitle("readExcel START"));
        stopWatch.start("read excel");

        Map<Integer, List<String>> data = new HashMap<>();

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(fileLocation))) {
            Sheet sheet = workbook.getSheetAt(0);

            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum > LIMIT_ROW) {
                throw new ExcelException("Excel max row is " + LIMIT_ROW + " row.");
            }

            short lastCellNum = sheet.getRow(0).getLastCellNum();
            log.info("# ==> lastCellNum = {}", lastCellNum);
            log.info("# ==> lastRowNum = {}", lastRowNum);

            int startRow = hasHeader ? 1 : 0;
            for (int currentRowNum = startRow; currentRowNum < lastRowNum + 1; currentRowNum++) {
                data.put(currentRowNum, new ArrayList<>());

                Row row = sheet.getRow(currentRowNum);
                rowToMap(data, lastCellNum, currentRowNum, row);
            }

        } catch (IOException e) {
            throw new ExcelException("Excel processing error.", e);

        } finally {
            stopWatch.stop();
            log.debug(logSubTitle("readExcel END"));
        }

        return data;
    }

    public static int readExcelAndFunction(String fileLocation, boolean hasHeader, StopWatch stopWatch,
            int functionBatchSize, ToIntFunction<Map<Integer, List<String>>> function) {
        log.debug(logSubTitle("readExcelAndFunction START"));
        stopWatch.start("read excel and function");

        int result;
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(fileLocation))) {
            Sheet sheet = workbook.getSheetAt(0);

            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum > LIMIT_ROW) {
                throw new ExcelException("Excel max row is " + LIMIT_ROW + " row.");
            }

            short lastCellNum = sheet.getRow(0).getLastCellNum();
            log.debug("# ==> lastCellNum = {}", lastCellNum);
            log.debug("# ==> lastRowNum = {}", lastRowNum);

            int processed = 0;
            int startRow = hasHeader ? 1 : 0;
            Map<Integer, List<String>> data = new HashMap<>();

            for (int currentRowNum = startRow; currentRowNum < lastRowNum + 1; currentRowNum++) {

                if (currentRowNum != 0 && currentRowNum % functionBatchSize == 0) {
                    // Apply function
                    processed += function.applyAsInt(data);
                    // clear Map
                    data.clear();
                }

                // create row in Map
                data.put(currentRowNum, new ArrayList<>());

                // save row data into Map
                Row row = sheet.getRow(currentRowNum);
                rowToMap(data, lastCellNum, currentRowNum, row);
            }

            // Apply function
            processed += function.applyAsInt(data);
            // clear Map
            data.clear();

            result = processed;

        } catch (IOException e) {
            throw new ExcelException("Excel processing error.", e);

        } finally {
            stopWatch.stop();
            log.debug(logSubTitle("readExcelAndFunction END"));
        }

        return result;
    }

    private static void rowToMap(Map<Integer, List<String>> data, short lastCellNum, int currentRowNum, Row row) {
        for (int currCellNum = 0; currCellNum < lastCellNum; currCellNum++) {
            Cell cell = row.getCell(currCellNum);

            if (cell == null) {
                data.get(currentRowNum).add("");
                continue;
            }

            switch (cell.getCellType()) {
                case STRING -> data.get(currentRowNum).add(cell.getRichStringCellValue().getString().trim());
                case NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        data.get(currentRowNum).add(cell.getLocalDateTimeCellValue() + "");
                    } else {
                        data.get(currentRowNum).add((long)cell.getNumericCellValue() + "");
                    }
                }
                case BOOLEAN -> data.get(currentRowNum).add(cell.getBooleanCellValue() + "");
                case FORMULA -> data.get(currentRowNum).add(cell.getCellFormula() + "");
                default -> data.get(currentRowNum).add("");
            }
        }
    }
}
