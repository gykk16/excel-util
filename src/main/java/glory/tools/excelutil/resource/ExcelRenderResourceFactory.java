package glory.tools.excelutil.resource;


import static org.apache.commons.lang3.reflect.FieldUtils.getAllFields;
import static org.springframework.core.annotation.AnnotationUtils.getAnnotation;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import glory.tools.excelutil.annotation.ExcelColumn;
import glory.tools.excelutil.annotation.ExcelColumnStyle;
import glory.tools.excelutil.annotation.ExcelDefaultBodyStyle;
import glory.tools.excelutil.annotation.ExcelDefaultHeaderStyle;
import glory.tools.excelutil.exception.InvalidExcelCellStyleException;
import glory.tools.excelutil.exception.NoExcelColumnAnnotationsException;
import glory.tools.excelutil.resource.collection.PreCalculatedCellStyleMap;
import glory.tools.excelutil.style.ExcelCellStyle;
import glory.tools.excelutil.style.NoExcelCellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.StringUtils;

/**
 * ExcelRenderResourceFactory
 */
public final class ExcelRenderResourceFactory {

    private ExcelRenderResourceFactory() {
    }

    public static ExcelRenderResource prepareRenderResource(Class<?> type, Workbook wb,
            DataFormatDecider dataFormatDecider) {

        PreCalculatedCellStyleMap styleMap = new PreCalculatedCellStyleMap(dataFormatDecider);
        Map<String, String> headerNamesMap = new LinkedHashMap<>();
        Map<String, Integer> headerWidthMap = new LinkedHashMap<>();
        List<String> fieldNames = new ArrayList<>();
        String countColumnName = "";

        ExcelColumnStyle classDefinedHeaderStyle = getHeaderExcelColumnStyle(type);
        ExcelColumnStyle classDefinedBodyStyle = getBodyExcelColumnStyle(type);

        for (Field field : getAllFields(type)) {
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);

                styleMap.put(
                        String.class,
                        ExcelCellKey.of(field.getName(), ExcelRenderLocation.HEADER),
                        getCellStyle(decideAppliedStyleAnnotation(classDefinedHeaderStyle, annotation.headerStyle())),
                        wb);

                Class<?> fieldType = field.getType();
                styleMap.put(
                        fieldType,
                        ExcelCellKey.of(field.getName(), ExcelRenderLocation.BODY),
                        getCellStyle(decideAppliedStyleAnnotation(classDefinedBodyStyle, annotation.bodyStyle())),
                        wb);

                fieldNames.add(field.getName());
                headerNamesMap.put(field.getName(),
                        StringUtils.hasText(annotation.headerName()) ? annotation.headerName() : field.getName());
                headerWidthMap.put(field.getName(), annotation.columnWidth());

                countColumnName = annotation.countColumn() ? field.getName() : countColumnName;
            }
            // dto 의 모든 필드 어노테이션 강제
            //            else {
            //                throw new ExcelException("All fields should be annotated with ExcelColumn annotation");
            //            }
        }

        if (styleMap.isEmpty()) {
            throw new NoExcelColumnAnnotationsException(String.format("Class %s has not @ExcelColumn at all", type));
        }
        return new ExcelRenderResource(styleMap, headerWidthMap, headerNamesMap, fieldNames, countColumnName);
    }

    private static ExcelColumnStyle getHeaderExcelColumnStyle(Class<?> clazz) {
        Annotation annotation = getAnnotation(clazz, ExcelDefaultHeaderStyle.class);
        if (annotation == null) {
            return null;
        }
        return ((ExcelDefaultHeaderStyle)annotation).style();
    }

    private static ExcelColumnStyle getBodyExcelColumnStyle(Class<?> clazz) {
        Annotation annotation = getAnnotation(clazz, ExcelDefaultBodyStyle.class);
        if (annotation == null) {
            return null;
        }
        return ((ExcelDefaultBodyStyle)annotation).style();
    }

    private static ExcelColumnStyle decideAppliedStyleAnnotation(
            ExcelColumnStyle classAnnotation,
            ExcelColumnStyle fieldAnnotation) {

        if (fieldAnnotation.excelCellStyleClass().equals(NoExcelCellStyle.class) && classAnnotation != null) {
            return classAnnotation;
        }
        return fieldAnnotation;
    }

    private static ExcelCellStyle getCellStyle(ExcelColumnStyle excelColumnStyle) {
        Class<? extends ExcelCellStyle> excelCellStyleClass = excelColumnStyle.excelCellStyleClass();
        // 1. Case of Enum
        if (excelCellStyleClass.isEnum()) {
            String enumName = excelColumnStyle.enumName();
            return findExcelCellStyle(excelCellStyleClass, enumName);
        }

        // 2. Case of Class
        try {
            return excelCellStyleClass.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | IllegalAccessException | NoSuchMethodException e) {
            throw new InvalidExcelCellStyleException(e.getMessage(), e);
        } catch (InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    @SuppressWarnings("unchecked")
    private static ExcelCellStyle findExcelCellStyle(Class<?> excelCellStyles, String enumName) {
        try {
            return (ExcelCellStyle)Enum.valueOf((Class<Enum>)excelCellStyles, enumName);
        } catch (NullPointerException e) {
            throw new InvalidExcelCellStyleException("enumName must not be null", e);
        } catch (IllegalArgumentException e) {
            throw new InvalidExcelCellStyleException(
                    "Enum %s does not name %s".formatted(excelCellStyles.getName(), enumName), e);
        }
    }

}
