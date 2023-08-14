package glory.tools.excelutil.resource;

import java.util.Objects;

import org.springframework.util.Assert;

public final class ExcelCellKey {

    private final String              dataFieldName;
    private final ExcelRenderLocation excelRenderLocation;

    private ExcelCellKey(String dataFieldName, ExcelRenderLocation excelRenderLocation) {
        this.dataFieldName = dataFieldName;
        this.excelRenderLocation = excelRenderLocation;
    }

    public static ExcelCellKey of(String fieldName, ExcelRenderLocation excelRenderLocation) {
        Assert.notNull(excelRenderLocation, "excelRenderLocation can not be null");
        return new ExcelCellKey(fieldName, excelRenderLocation);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        ExcelCellKey that = (ExcelCellKey)o;
        return Objects.equals(dataFieldName, that.dataFieldName) &&
               excelRenderLocation == that.excelRenderLocation;
    }

    @Override
    public int hashCode() {
        return Objects.hash(dataFieldName, excelRenderLocation);
    }

}
