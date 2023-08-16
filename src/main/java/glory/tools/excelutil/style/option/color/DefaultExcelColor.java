package glory.tools.excelutil.style.option.color;

import glory.tools.excelutil.exception.UnSupportedExcelTypeException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class DefaultExcelColor implements ExcelColor {

    private static final int MIN_RGB = 0;
    private static final int MAX_RGB = 255;

    private final byte red;
    private final byte green;
    private final byte blue;

    private DefaultExcelColor(byte red, byte green, byte blue) {
        this.red = red;
        this.green = green;
        this.blue = blue;
    }

    public static DefaultExcelColor rgb(int red, int green, int blue) {
        validateRGB(red, green, blue);
        return new DefaultExcelColor((byte)red, (byte)green, (byte)blue);
    }

    /**
     * applyForeground In current, only supports XSSFCellStyle because can not find HSSFCellStyle RGB configuration
     * Please share if you find the HSSFCellStyle RGB configuration
     */
    @Override
    public void applyForeground(CellStyle cellStyle) {
        if (!(cellStyle instanceof XSSFCellStyle xssfCellStyle)) {
            throw new UnSupportedExcelTypeException(
                    "Unsupported Excel Type: %s. Only XSSFCellStyle is supported."
                            .formatted(cellStyle.getClass().getName()));
        }

        xssfCellStyle.setFillForegroundColor(
                new XSSFColor(new byte[] {red, green, blue}, new DefaultIndexedColorMap()));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private static void validateRGB(int red, int green, int blue) {
        if (red < MIN_RGB || red > MAX_RGB
            || green < MIN_RGB || green > MAX_RGB
            || blue < MIN_RGB || blue > MAX_RGB) {
            throw new IllegalArgumentException("Wrong RGB(%s %s %s)".formatted(red, green, blue));
        }
    }

}
