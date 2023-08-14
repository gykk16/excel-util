package glory.tools.excelutil.style.configurer;

import glory.tools.excelutil.style.option.align.ExcelAlign;
import glory.tools.excelutil.style.option.align.NoExcelAlign;
import glory.tools.excelutil.style.option.border.ExcelBorders;
import glory.tools.excelutil.style.option.border.NoExcelBorders;
import glory.tools.excelutil.style.option.color.DefaultExcelColor;
import glory.tools.excelutil.style.option.color.ExcelColor;
import glory.tools.excelutil.style.option.color.NoExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelCellStyleConfigurer {

    private ExcelAlign   excelAlign      = new NoExcelAlign();
    private ExcelColor   foregroundColor = new NoExcelColor();
    private ExcelBorders excelBorders    = new NoExcelBorders();

    public ExcelCellStyleConfigurer() {
    }

    public ExcelCellStyleConfigurer excelAlign(ExcelAlign excelAlign) {
        this.excelAlign = excelAlign;
        return this;
    }

    public ExcelCellStyleConfigurer foregroundColor(int red, int blue, int green) {
        this.foregroundColor = DefaultExcelColor.rgb(red, blue, green);
        return this;
    }

    public ExcelCellStyleConfigurer excelBorders(ExcelBorders excelBorders) {
        this.excelBorders = excelBorders;
        return this;
    }

    public void configure(CellStyle cellStyle) {
        excelAlign.apply(cellStyle);
        foregroundColor.applyForeground(cellStyle);
        excelBorders.apply(cellStyle);
    }

}
