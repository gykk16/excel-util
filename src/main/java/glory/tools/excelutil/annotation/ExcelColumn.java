package glory.tools.excelutil.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import glory.tools.excelutil.style.NoExcelCellStyle;
import org.springframework.core.annotation.AliasFor;

/**
 * ExcelColumn Annotation
 * <p>
 * annotation for excel column
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    /**
     * column name
     */
    String headerName() default "";

    /**
     * column header style
     */
    ExcelColumnStyle headerStyle() default @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class);

    /**
     * column body style
     */
    ExcelColumnStyle bodyStyle() default @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class);

    /**
     * column width
     */
    int columnWidth() default 20;

    /**
     * true if it is count column
     * <p>
     * count column is column that count row number automatically
     */
    boolean countColumn() default false;

}
