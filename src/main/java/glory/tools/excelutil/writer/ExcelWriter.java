package glory.tools.excelutil.writer;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public interface ExcelWriter<T> {

    void write(OutputStream stream) throws IOException;

    void addRows(List<T> data);

}
