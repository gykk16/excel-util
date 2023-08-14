package glory.tools.excelutil.annotation;


import glory.tools.excelutil.style.CustomExcelCellStyle;
import glory.tools.excelutil.style.DefaultExcelCellStyle;
import glory.tools.excelutil.style.ExcelCellStyle;

public @interface ExcelColumnStyle {

    /**
     * Enum implements {@link ExcelCellStyle} Also, can use just class. If not use Enum, enumName will be ignored
     *
     * @see DefaultExcelCellStyle
     * @see CustomExcelCellStyle
     */
    Class<? extends ExcelCellStyle> excelCellStyleClass();

    /**
     * name of Enum implements {@link ExcelCellStyle} if not use Enum, enumName will be ignored
     */
    String enumName() default "";

}
