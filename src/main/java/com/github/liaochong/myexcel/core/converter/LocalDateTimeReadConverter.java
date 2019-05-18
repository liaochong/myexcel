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
package com.github.liaochong.myexcel.core.converter;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.utils.StringUtil;

import java.lang.reflect.Field;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Objects;
import java.util.TimeZone;
import java.util.regex.Pattern;

/**
 * LocalDateTime读取转换器
 *
 * @author liaochong
 * @version 1.0
 */
public class LocalDateTimeReadConverter implements Converter<String, LocalDateTime> {

    private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    private static final Pattern NUMBER_PATTERN = Pattern.compile("^\\d+$");

    @Override
    public LocalDateTime convert(String obj, Field field) {
        if (StringUtil.isBlank(obj)) {
            return null;
        }
        String trimContent = obj.trim();
        boolean isNumber = NUMBER_PATTERN.matcher(trimContent).find();
        if (isNumber) {
            final long time = Long.parseLong(trimContent);
            return LocalDateTime.ofInstant(Instant.ofEpochSecond(time), TimeZone
                    .getDefault().toZoneId());
        }
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        String dateFormat = DEFAULT_DATE_FORMAT;
        if (Objects.nonNull(excelColumn) && StringUtil.isNotBlank(excelColumn.dateFormatPattern())) {
            dateFormat = excelColumn.dateFormatPattern();
        }
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(dateFormat);
        return LocalDateTime.parse(trimContent, dateTimeFormatter);
    }
}
