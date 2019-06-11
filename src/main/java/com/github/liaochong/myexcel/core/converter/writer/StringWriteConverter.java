package com.github.liaochong.myexcel.core.converter.writer;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.WriteConverter;

import java.lang.reflect.Field;
import java.util.Objects;

/**
 * @author simon
 * @version 2.5.0
 */
public class StringWriteConverter implements WriteConverter {

    @Override
    public boolean support(Field field, Object fieldVal) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        return Objects.nonNull(excelColumn) && Objects.nonNull(fieldVal) && excelColumn.convertToString();
    }

    @Override
    public Pair<Class, Object> convert(Field field, Object fieldVal) {
        return Pair.of(String.class, fieldVal.toString());
    }
}
