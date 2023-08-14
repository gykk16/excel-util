package glory.tools.excelutil.resource;

import java.util.List;
import java.util.Map;

import glory.tools.excelutil.resource.collection.PreCalculatedCellStyleMap;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelRenderResource {

    private final PreCalculatedCellStyleMap styleMap;

    // TODO dataFieldName -> excelHeaderName Map Abstraction
    private final Map<String, String>  excelHeaderNames;
    private final Map<String, Integer> excelHeaderWidth;
    private final List<String>         dataFieldNames;
    private final String               countColumnName;

    public ExcelRenderResource(PreCalculatedCellStyleMap styleMap, Map<String, Integer> excelHeaderWidth,
            Map<String, String> excelHeaderNames, List<String> dataFieldNames, String countColumnName) {
        this.styleMap = styleMap;
        this.excelHeaderWidth = excelHeaderWidth;
        this.excelHeaderNames = excelHeaderNames;
        this.dataFieldNames = dataFieldNames;
        this.countColumnName = countColumnName;
    }

    public CellStyle getCellStyle(String dataFieldName, ExcelRenderLocation excelRenderLocation) {
        return styleMap.get(ExcelCellKey.of(dataFieldName, excelRenderLocation));
    }

    public String getExcelHeaderName(String dataFieldName) {
        return excelHeaderNames.get(dataFieldName);
    }

    public int getExcelHeaderWidth(String dataFieldName) {
        return excelHeaderWidth.get(dataFieldName);
    }

    public List<String> getDataFieldNames() {
        return dataFieldNames;
    }

    public String getCountColumnName() {
        return countColumnName;
    }
}
