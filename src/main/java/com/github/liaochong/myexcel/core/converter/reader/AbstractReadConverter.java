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
package com.github.liaochong.myexcel.core.converter.reader;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.converter.Converter;
import com.github.liaochong.myexcel.utils.StringUtil;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.Objects;
import java.util.regex.Pattern;

/**
 * 读取转换抽象
 *
 * @author chd.y
 * @since 2.3.1
 */
public abstract class AbstractReadConverter<R> implements Converter<String, R> {

    protected static WeakCache<String, DateTimeFormatter> dateTimeFormatterWeakCache = new WeakCache<>();

    protected static WeakCache<String, SimpleDateFormat> simpleDateFormatWeakCache = new WeakCache<>();

    private static final Pattern PATTERN_NUMBER = Pattern.compile("^\\d+$");

    protected static final Pattern PATTERN_NON_NUMBER = Pattern.compile("[^\\.\\d\\-]");

    protected static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

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
     *
     * @param v     待转换值
     * @param field 待转换值所属字段
     * @return 目标值
     */
    protected abstract R doConvert(String v, Field field);


    /**
     * 取得DateFormatPattern
     *
     * @param field 待转换值所属字段
     * @return 时间格式
     */
    protected String getDateFormatPattern(Field field) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        if (Objects.nonNull(excelColumn) && StringUtil.isNotBlank(excelColumn.dateFormatPattern())) {
            return excelColumn.dateFormatPattern();
        }
        return DEFAULT_DATE_FORMAT;
    }

    /**
     * 是否为数值
     *
     * @param v 内容
     * @return true/false
     */
    protected boolean isNumber(String v) {
        return PATTERN_NUMBER.matcher(v).matches();
    }

    /**
     * 获取DateTimeFormatter
     *
     * @param field 字段
     * @return DateTimeFormatter
     */
    protected DateTimeFormatter getDateFormatFormatter(Field field) {
        String dateFormatPattern = getDateFormatPattern(field);
        DateTimeFormatter dateTimeFormatter = dateTimeFormatterWeakCache.get(dateFormatPattern);
        if (Objects.isNull(dateTimeFormatter)) {
            dateTimeFormatter = DateTimeFormatter.ofPattern(dateFormatPattern);
            dateTimeFormatterWeakCache.cache(dateFormatPattern, dateTimeFormatter);
        }
        return dateTimeFormatter;
    }
}
