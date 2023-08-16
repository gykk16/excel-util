package glory.tools.excelutil.style.option.border;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;

public final class DefaultExcelBorders implements ExcelBorders {

    private final List<? extends ExcelBorder> borders;

    public DefaultExcelBorders(List<? extends ExcelBorder> borders) {
        validateBorders(borders);
        this.borders = borders;
    }

    public static DefaultExcelBorders newInstance(ExcelBorderStyle style) {
        List<DefaultExcelBorder> excelBorders = new ArrayList<>();
        for (int i = 0; i < 4; i++) {
            excelBorders.add(new DefaultExcelBorder(style));
        }
        return new DefaultExcelBorders(excelBorders);
    }

    public void apply(CellStyle cellStyle) {
        borders.get(0).applyTop(cellStyle);
        borders.get(1).applyRight(cellStyle);
        borders.get(2).applyBottom(cellStyle);
        borders.get(3).applyLeft(cellStyle);
    }

    private void validateBorders(List<? extends ExcelBorder> borders) {
        if (borders.size() != 4) {
            throw new IllegalArgumentException("Should be initialized with TOP RIGHT LEFT BOTTOM borders");
        }
    }

}
