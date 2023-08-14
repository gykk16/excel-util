package glory.tools.excelutil.writer;

import glory.tools.excelutil.annotation.ExcelColumn;
import glory.tools.excelutil.annotation.ExcelDefaultBodyStyle;
import glory.tools.excelutil.annotation.ExcelDefaultHeaderStyle;

@ExcelDefaultHeaderStyle
@ExcelDefaultBodyStyle
public class TestDataDTO {

    @ExcelColumn(headerName = "no", countColumn = true)
    private int count;

    @ExcelColumn(headerName = "name")
    private String name;

    @ExcelColumn(headerName = "age")
    private int age;

    public TestDataDTO(String name, int age) {
        this.name = name;
        this.age = age;
    }

    public int getCount() {
        return count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }
}
