package glory.tools.excelutil.style;

import glory.tools.excelutil.style.configurer.ExcelCellStyleConfigurer;
import org.apache.poi.ss.usermodel.CellStyle;

public abstract class CustomExcelCellStyle implements ExcelCellStyle {

    private final ExcelCellStyleConfigurer configurer = new ExcelCellStyleConfigurer();

    protected CustomExcelCellStyle() {
        configure(configurer);
    }

    @Override
    public void apply(CellStyle cellStyle) {
        configurer.configure(cellStyle);
    }

    public abstract void configure(ExcelCellStyleConfigurer configurer);

}
