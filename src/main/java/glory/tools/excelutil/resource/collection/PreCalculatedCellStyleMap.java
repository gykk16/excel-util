package glory.tools.excelutil.resource.collection;


import java.util.HashMap;
import java.util.Map;

import glory.tools.excelutil.resource.DataFormatDecider;
import glory.tools.excelutil.resource.ExcelCellKey;
import glory.tools.excelutil.style.ExcelCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * PreCalculatedCellStyleMap
 * <p>
 * Determines cell's style In currently, PreCalculatedCellStyleMap determines {org.apache.poi.ss.usermodel.DataFormat}
 */
public class PreCalculatedCellStyleMap {

    private final DataFormatDecider            dataFormatDecider;
    private final Map<ExcelCellKey, CellStyle> cellStyleMap = new HashMap<>();


    public PreCalculatedCellStyleMap(DataFormatDecider dataFormatDecider) {
        this.dataFormatDecider = dataFormatDecider;
    }

    public void put(Class<?> fieldType, ExcelCellKey excelCellKey, ExcelCellStyle excelCellStyle, Workbook wb) {
        CellStyle cellStyle = wb.createCellStyle();
        DataFormat dataFormat = wb.createDataFormat();
        Font font = wb.createFont();
        font.setFontHeightInPoints((short)9);
        cellStyle.setDataFormat(dataFormatDecider.getDataFormat(wb, dataFormat, fieldType));
        cellStyle.setFont(font);
        excelCellStyle.apply(cellStyle);
        cellStyleMap.put(excelCellKey, cellStyle);
    }

    public CellStyle get(ExcelCellKey excelCellKey) {
        return cellStyleMap.get(excelCellKey);
    }

    public boolean isEmpty() {
        return cellStyleMap.isEmpty();
    }

}
