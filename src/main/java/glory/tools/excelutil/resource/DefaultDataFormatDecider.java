package glory.tools.excelutil.resource;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

public class DefaultDataFormatDecider implements DataFormatDecider {

    private static final String CURRENT_FORMAT                = "#,##0";
    private static final String FLOAT_FORMAT_2_DECIMAL_PLACES = "#,##0.00";
    private static final String DEFAULT_FORMAT                = "";

    @Override
    public short getDataFormat(DataFormat dataFormat, Class<?> type) {
        if (isFloatType(type)) {
            return dataFormat.getFormat(FLOAT_FORMAT_2_DECIMAL_PLACES);
        }

        if (isIntegerType(type)) {
            return dataFormat.getFormat(CURRENT_FORMAT);
        }
        return dataFormat.getFormat(DEFAULT_FORMAT);
    }

    @Override
    public short getDataFormat(Workbook workbook, DataFormat dataFormat, Class<?> type) {
        if (isFloatType(type)) {
            return dataFormat.getFormat(FLOAT_FORMAT_2_DECIMAL_PLACES);
        }

        if (isIntegerType(type)) {
            return dataFormat.getFormat(CURRENT_FORMAT);
        }

        if (isDateType(type)) {
            CreationHelper creationHelper = workbook.getCreationHelper();
            if (type == LocalDate.class) {
                return creationHelper.createDataFormat().getFormat("yyyy-mm-dd");
            } else {
                return creationHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss");
            }
        }
        return dataFormat.getFormat(DEFAULT_FORMAT);
    }

    private boolean isFloatType(Class<?> type) {
        List<Class<?>> floatTypes = Arrays.asList(
                Float.class, float.class,
                Double.class, double.class
        );
        return floatTypes.contains(type);
    }

    private boolean isIntegerType(Class<?> type) {
        List<Class<?>> integerTypes = Arrays.asList(
                Byte.class, byte.class,
                Short.class, short.class,
                Integer.class, int.class,
                Long.class, long.class
        );
        return integerTypes.contains(type);
    }

    private boolean isDateType(Class<?> type) {
        List<Class<?>> dateTypes = Arrays.asList(
                LocalDate.class,
                LocalDateTime.class
        );
        return dateTypes.contains(type);
    }

}
