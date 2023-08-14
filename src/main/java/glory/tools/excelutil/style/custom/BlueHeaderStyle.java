package glory.tools.excelutil.style.custom;


import glory.tools.excelutil.style.CustomExcelCellStyle;
import glory.tools.excelutil.style.configurer.ExcelCellStyleConfigurer;
import glory.tools.excelutil.style.option.align.DefaultExcelAlign;
import glory.tools.excelutil.style.option.border.DefaultExcelBorders;
import glory.tools.excelutil.style.option.border.ExcelBorderStyle;

public class BlueHeaderStyle extends CustomExcelCellStyle {

    @Override
    public void configure(ExcelCellStyleConfigurer configurer) {
        configurer.foregroundColor(223, 235, 246)
                .excelBorders(DefaultExcelBorders.newInstance(ExcelBorderStyle.THIN))
                .excelAlign(DefaultExcelAlign.CENTER_CENTER);
    }

}