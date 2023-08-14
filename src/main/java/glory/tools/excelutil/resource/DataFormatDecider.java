package glory.tools.excelutil.resource;

import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

public interface DataFormatDecider {

    short getDataFormat(DataFormat dataFormat, Class<?> type);

    short getDataFormat(Workbook workbook, DataFormat dataFormat, Class<?> type);

}
