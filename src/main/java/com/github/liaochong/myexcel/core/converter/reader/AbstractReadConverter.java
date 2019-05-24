package com.github.liaochong.myexcel.core.converter.reader;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.converter.Converter;
import com.github.liaochong.myexcel.utils.StringUtil;

import java.lang.reflect.Field;
import java.util.Objects;

/**
 * 读取转换抽象
 * @author chd.y
 * @param <R>
 */
public abstract class AbstractReadConverter<R> implements Converter<String, R> {

    @Override
    public R convert(String obj, Field field) {
        if (StringUtil.isBlank(obj)) {
            return null;
        }
        String trimContent = obj.trim();
        return doConvert(trimContent, field);
    }


    /**
     * 把输入参数进行处理后，转换。
     * @param v
     * @param field
     * @return
     */
    protected abstract R doConvert(String v, Field field);


    /**
     * 取得DateFormatPattern
     * @param field
     * @param defaultDateFormat
     * @return
     */
    protected String getDateFormatPattern(Field field, String defaultDateFormat) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        String dateFormat = defaultDateFormat;
        if (Objects.nonNull(excelColumn) && StringUtil.isNotBlank(excelColumn.dateFormatPattern())) {
            dateFormat = excelColumn.dateFormatPattern();
        }
        return dateFormat;
    }
}
