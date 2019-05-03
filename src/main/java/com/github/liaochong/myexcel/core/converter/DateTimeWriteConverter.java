package com.github.liaochong.myexcel.core.converter;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.cache.Cache;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.utils.StringUtil;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public class DateTimeWriteConverter implements WriteConverter {

    private static final Cache<String, DateTimeFormatter> DATETIME_FORMATTER_CONTAINER = new WeakCache<>();

    @Override
    public Pair<Class, Object> convert(Field field, Object fieldVal) {
        Class<?> fieldType = field.getType();
        if (fieldType != LocalDateTime.class && fieldType != LocalDate.class && fieldType != Date.class) {
            return new Pair<>(fieldType, fieldVal);
        }
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        if (Objects.isNull(excelColumn) || Objects.isNull(fieldVal)) {
            return new Pair<>(fieldType, fieldVal);
        }
        // 时间格式化
        String dateFormatPattern = excelColumn.dateFormatPattern();
        if (StringUtil.isBlank(dateFormatPattern)) {
            return new Pair<>(fieldType, fieldVal);
        }
        if (fieldType == LocalDateTime.class) {
            LocalDateTime localDateTime = (LocalDateTime) fieldVal;
            DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
            return new Pair<>(String.class, formatter.format(localDateTime));
        } else if (fieldType == LocalDate.class) {
            LocalDate localDate = (LocalDate) fieldVal;
            DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
            return new Pair<>(String.class, formatter.format(localDate));
        }
        Date date = (Date) fieldVal;
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormatPattern);
        return new Pair<>(String.class, simpleDateFormat.format(date));
    }

    /**
     * 获取时间格式化
     *
     * @param dateFormat 时间格式化
     * @return DateTimeFormatter
     */
    private DateTimeFormatter getDateTimeFormatter(String dateFormat) {
        DateTimeFormatter formatter = DATETIME_FORMATTER_CONTAINER.get(dateFormat);
        if (Objects.isNull(formatter)) {
            formatter = DateTimeFormatter.ofPattern(dateFormat);
            DATETIME_FORMATTER_CONTAINER.cache(dateFormat, formatter);
        }
        return formatter;
    }
}
