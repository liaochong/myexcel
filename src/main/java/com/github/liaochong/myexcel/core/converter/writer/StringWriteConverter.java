package com.github.liaochong.myexcel.core.converter.writer;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.WriteConverter;

import java.lang.reflect.Field;
import java.util.Objects;

/**
 * @author simon
 * @version 1.0
 * @date 2019-06-06 15:52
 */
public class StringWriteConverter implements WriteConverter {

    @Override
    public boolean support(Field field, Object fieldVal) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        return Objects.nonNull(excelColumn) && Objects.nonNull(fieldVal);
    }

    @Override
    public Pair<Class, Object> convert(Field field, Object fieldVal) {
        Class<?> fieldType = field.getType();
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        boolean convertToString = excelColumn.convertToString();
        return convertToString ? Pair.of(String.class, fieldVal.toString()) : Pair.of(fieldType, fieldVal);
    }
}
