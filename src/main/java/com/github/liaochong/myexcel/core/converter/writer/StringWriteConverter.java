package com.github.liaochong.myexcel.core.converter.writer;

import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.ConvertContext;
import com.github.liaochong.myexcel.core.converter.WriteConverter;

import java.lang.reflect.Field;

/**
 * @author simon
 * @version 2.5.0
 */
public class StringWriteConverter implements WriteConverter {

    @Override
    public boolean support(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        ExcelColumnMapping mapping = convertContext.getExcelColumnMappingMap().get(field);
        return mapping != null && mapping.isConvertToString();
    }

    @Override
    public Pair<Class, Object> convert(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        return Pair.of(String.class, fieldVal.toString());
    }
}
