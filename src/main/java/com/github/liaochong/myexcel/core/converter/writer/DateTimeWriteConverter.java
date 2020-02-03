/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core.converter.writer;

import com.github.liaochong.myexcel.core.ConvertContext;
import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.cache.Cache;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.WriteConverter;
import com.github.liaochong.myexcel.utils.ReflectUtil;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * @author liaochong
 * @version 1.0
 */
public class DateTimeWriteConverter implements WriteConverter {

    private static final Cache<String, DateTimeFormatter> DATETIME_FORMATTER_CONTAINER = new WeakCache<>();

    private static final WeakCache<String, ThreadLocal<SimpleDateFormat>> SIMPLE_DATE_FORMAT_WEAK_CACHE = new WeakCache<>();

    @Override
    public boolean support(Field field, Object fieldVal, ConvertContext convertContext) {
        return ReflectUtil.isDate(field.getType());
    }

    @Override
    public Pair<Class, Object> convert(Field field, Object fieldVal, ConvertContext convertContext) {
        Class<?> fieldType = field.getType();
        // 时间格式化
        String dateFormatPattern = getDateFormatPattern(convertContext, field);
        if (fieldType == LocalDateTime.class) {
            LocalDateTime localDateTime = (LocalDateTime) fieldVal;
            DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
            return Pair.of(String.class, formatter.format(localDateTime));
        } else if (fieldType == LocalDate.class) {
            LocalDate localDate = (LocalDate) fieldVal;
            DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
            return Pair.of(String.class, formatter.format(localDate));
        }
        SimpleDateFormat simpleDateFormat = this.getSimpleDateFormat(dateFormatPattern);
        return Pair.of(String.class, simpleDateFormat.format((Date) fieldVal));
    }

    private String getDateFormatPattern(ConvertContext convertContext, Field field) {
        ExcelColumnMapping mapping = convertContext.getExcelColumnMappingMap().get(field);
        if (mapping == null) {
            return field.getType() == LocalDate.class ? convertContext.getGlobalSetting().getDateFormat() : convertContext.getGlobalSetting().getDateTimeFormat();
        }
        String dateFormatPattern = mapping.getFormat();
        if (dateFormatPattern.isEmpty()) {
            dateFormatPattern = field.getType() == LocalDate.class ? convertContext.getGlobalSetting().getDateFormat() : convertContext.getGlobalSetting().getDateTimeFormat();
        }
        return dateFormatPattern;
    }

    /**
     * 获取时间格式化
     *
     * @param dateFormat 时间格式化
     * @return DateTimeFormatter
     */
    private DateTimeFormatter getDateTimeFormatter(String dateFormat) {
        DateTimeFormatter formatter = DATETIME_FORMATTER_CONTAINER.get(dateFormat);
        if (formatter == null) {
            formatter = DateTimeFormatter.ofPattern(dateFormat);
            DATETIME_FORMATTER_CONTAINER.cache(dateFormat, formatter);
        }
        return formatter;
    }

    private SimpleDateFormat getSimpleDateFormat(String dateFormatPattern) {
        ThreadLocal<SimpleDateFormat> tl = SIMPLE_DATE_FORMAT_WEAK_CACHE.get(dateFormatPattern);
        if (tl == null) {
            tl = ThreadLocal.withInitial(() -> new SimpleDateFormat(dateFormatPattern));
            SIMPLE_DATE_FORMAT_WEAK_CACHE.cache(dateFormatPattern, tl);
        }
        return tl.get();
    }
}
