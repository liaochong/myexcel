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

import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.converter.ConvertContext;
import com.github.liaochong.myexcel.core.converter.ReadConverter;
import com.github.liaochong.myexcel.utils.StringUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Objects;
import java.util.regex.Pattern;

/**
 * 读取转换抽象
 *
 * @author chd.y
 * @since 2.3.1
 */
public abstract class AbstractReadConverter<R> implements ReadConverter<String, R> {

    protected static final WeakCache<String, DateTimeFormatter> DATE_TIME_FORMATTER_WEAK_CACHE = new WeakCache<>();

    protected static final WeakCache<String, ThreadLocal<SimpleDateFormat>> SIMPLE_DATE_FORMAT_WEAK_CACHE = new WeakCache<>();

    /**
     * 时间数字正则表达式
     */
    private static final Pattern PATTERN_DATE_NUMBER = Pattern.compile("^[1-9]\\d{10,}$");

    /**
     * 数字、小数正则表达式
     */
    private static final Pattern PATTERN_DATE_DECIMAL = Pattern.compile("[0-9]+\\.*[0-9]*");

    private static final LocalDateTime START_LOCAL_DATE_TIME = LocalDateTime.of(1900, 1, 1, 0, 0, 0);

    @Override
    public R convert(String obj, Field field, ConvertContext convertContext) {
        if (StringUtil.isBlank(obj)) {
            return null;
        }
        String trimContent = obj.trim();
        // negative
        if (trimContent.startsWith(Constants.LEFT_BRACKET)) {
            if (trimContent.endsWith(Constants.RIGHT_BRACKET)) {
                int length = trimContent.length();
                if (length > 2) {
                    trimContent = "-" + trimContent.substring(1, length - 1);
                }
            }
        }
        return doConvert(trimContent, field, convertContext);
    }


    /**
     * 把输入参数进行处理后，转换。
     *
     * @param v              待转换值
     * @param field          待转换值所属字段
     * @param convertContext 转换上下文
     * @return 目标值
     */
    protected abstract R doConvert(String v, Field field, ConvertContext convertContext);


    /**
     * 取得DateFormatPattern
     *
     * @param field          待转换值所属字段
     * @param convertContext 转换上下文
     * @return 时间格式
     */
    protected String getDateFormatPattern(Field field, ConvertContext convertContext) {
        ExcelColumnMapping mapping = convertContext.getExcelColumnMappingMap().get(field);
        if (mapping == null) {
            return field.getType() == LocalDate.class ? convertContext.getConfiguration().getDateFormat() : field.getType() == LocalTime.class ? convertContext.getConfiguration().getLocalTimeFormat() : convertContext.getConfiguration().getDateTimeFormat();
        }
        String format = mapping.getFormat();
        if (!format.isEmpty()) {
            return format;
        }
        return field.getType() == LocalDate.class ? convertContext.getConfiguration().getDateFormat() : field.getType() == LocalTime.class ? convertContext.getConfiguration().getLocalTimeFormat() : convertContext.getConfiguration().getDateTimeFormat();
    }

    /**
     * 是否为时间类数值
     *
     * @param v 内容
     * @return true/false
     */
    protected boolean isDateNumber(String v) {
        return PATTERN_DATE_NUMBER.matcher(v).matches();
    }

    /**
     * 是否为Excel数字日期
     *
     * @param v 内容
     * @return true/false
     */
    protected boolean isDateDecimalNumber(String v) {
        return PATTERN_DATE_DECIMAL.matcher(v).matches();
    }

    /**
     * 获取DateTimeFormatter
     *
     * @param field          字段
     * @param convertContext 转换上下文
     * @return DateTimeFormatter
     */
    protected DateTimeFormatter getDateFormatFormatter(Field field, ConvertContext convertContext) {
        String dateFormatPattern = getDateFormatPattern(field, convertContext);
        DateTimeFormatter dateTimeFormatter = DATE_TIME_FORMATTER_WEAK_CACHE.get(dateFormatPattern);
        if (Objects.isNull(dateTimeFormatter)) {
            dateTimeFormatter = DateTimeFormatter.ofPattern(dateFormatPattern);
            DATE_TIME_FORMATTER_WEAK_CACHE.cache(dateFormatPattern, dateTimeFormatter);
        }
        return dateTimeFormatter;
    }

    protected SimpleDateFormat getSimpleDateFormat(String dateFormatPattern) {
        ThreadLocal<SimpleDateFormat> tl = SIMPLE_DATE_FORMAT_WEAK_CACHE.get(dateFormatPattern);
        if (tl == null) {
            tl = ThreadLocal.withInitial(() -> new SimpleDateFormat(dateFormatPattern));
            SIMPLE_DATE_FORMAT_WEAK_CACHE.cache(dateFormatPattern, tl);
        }
        return tl.get();
    }

    /**
     * 将Excel转换的数字日期转换为时间戳
     *
     * @param value 数字日期，例如43728.9319444444
     * @return 时间戳
     */
    protected long convertExcelNumberDateToMilli(String value) {
        //如果是数字 小于0则 返回
        BigDecimal bd = new BigDecimal(value);
        int days = bd.intValue();
        int seconds = (int) Math.round(bd.subtract(new BigDecimal(days)).doubleValue() * 24 * 3600);
        //获取时间
        int hour = seconds / 3600;
        int secondsOfHours = hour * 3600;
        int minute = (seconds - secondsOfHours) / 60;
        int second = seconds - secondsOfHours - minute * 60;
        LocalDateTime localDateTime = START_LOCAL_DATE_TIME.plusDays(days - 2).plusHours(hour).plusMinutes(minute).plusSeconds(second);
        return localDateTime.atZone(ZoneId.systemDefault()).toInstant().toEpochMilli();
    }
}
