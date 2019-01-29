package com.github.liaochong.html2excel.core.converter;

import com.github.liaochong.html2excel.core.annotation.ExcelColumn;
import com.github.liaochong.html2excel.core.cache.Cache;
import com.github.liaochong.html2excel.core.cache.WeakCache;
import com.github.liaochong.html2excel.utils.StringUtil;

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
public class DateTimeConverter implements Converter {

    private static final Cache<String, DateTimeFormatter> DATETIME_FORMATTER_CONTAINER = new WeakCache<>();

    @Override
    public Object convert(Field field, Object object) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        if (Objects.isNull(excelColumn) || Objects.isNull(object)) {
            return object;
        }
        // 时间格式化
        String dateFormatPattern = excelColumn.dateFormatPattern();
        if (StringUtil.isNotBlank(dateFormatPattern)) {
            Class<?> fieldType = field.getType();
            if (fieldType == LocalDateTime.class) {
                LocalDateTime localDateTime = (LocalDateTime) object;
                DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
                return formatter.format(localDateTime);
            } else if (fieldType == LocalDate.class) {
                LocalDate localDate = (LocalDate) object;
                DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
                return formatter.format(localDate);
            } else if (fieldType == Date.class) {
                Date date = (Date) object;
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormatPattern);
                return simpleDateFormat.format(date);
            }
        }
        return object;
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
