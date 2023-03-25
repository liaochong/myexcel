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
import com.github.liaochong.myexcel.core.context.ReadContext;
import com.github.liaochong.myexcel.core.converter.ReadConverter;
import com.github.liaochong.myexcel.utils.StringUtil;

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
public abstract class AbstractReadConverter<R> implements ReadConverter<R> {

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
    public R convert(ReadContext<?> readContext) {
        if (StringUtil.isBlank(readContext.getVal())) {
            return null;
        }
        String trimContent = readContext.getVal().trim();
        // negative
        if (trimContent.startsWith(Constants.LEFT_BRACKET)) {
            if (trimContent.endsWith(Constants.RIGHT_BRACKET)) {
                int length = trimContent.length();
                if (length > 2) {
                    trimContent = "-" + trimContent.substring(1, length - 1);
                }
            }
        }
        readContext.setVal(trimContent);
        return doConvert(readContext);
    }


    /**
     * 把输入参数进行处理后，转换。
     *
     * @param readContext 读取转换上下文
     * @return 目标值
     */
    protected abstract R doConvert(ReadContext<?> readContext);


    /**
     * 取得DateFormatPattern
     *
     * @param readContext 转换上下文
     * @return 时间格式
     */
    protected String getDateFormatPattern(ReadContext<?> readContext) {
        ExcelColumnMapping mapping = readContext.convertContext.excelColumnMappingMap.get(readContext.getField());
        if (mapping == null) {
            return readContext.getField().getType() == LocalDate.class ? readContext.convertContext.configuration.dateFormat : readContext.getField().getType() == LocalTime.class ? readContext.convertContext.configuration.localTimeFormat : readContext.convertContext.configuration.dateTimeFormat;
        }
        String format = mapping.format;
        if (!format.isEmpty()) {
            return format;
        }
        return readContext.getField().getType() == LocalDate.class ? readContext.convertContext.configuration.dateFormat : readContext.getField().getType() == LocalTime.class ? readContext.convertContext.configuration.localTimeFormat : readContext.convertContext.configuration.dateTimeFormat;
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
     * @param readContext 转换上下文
     * @return DateTimeFormatter
     */
    protected DateTimeFormatter getDateFormatFormatter(ReadContext<?> readContext) {
        String dateFormatPattern = getDateFormatPattern(readContext);
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
