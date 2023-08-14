package glory.tools.excelutil.style.option.color;

import org.apache.poi.ss.usermodel.CellStyle;

public class NoExcelColor implements ExcelColor {

    @Override
    public void applyForeground(CellStyle cellStyle) {
        // Do nothing
    }

}
