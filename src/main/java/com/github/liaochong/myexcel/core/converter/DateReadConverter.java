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
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Objects;
import java.util.TimeZone;
import java.util.regex.Pattern;

/**
 * @author liaochong
 * @version 1.0
 */
public class DateReadConverter implements ReadConverter {

    private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    private static final Pattern NUMBER_PATTERN = Pattern.compile("^\\d+$");

    @Override
    public boolean convert(String content, Field field, Object obj) throws Exception {
        if (StringUtil.isBlank(content)) {
            return false;
        }
        Class<?> type = field.getType();
        if (type != Date.class && type != LocalDate.class && type != LocalDateTime.class) {
            return false;
        }
        String trimContent = content.trim();
        boolean isNumber = NUMBER_PATTERN.matcher(trimContent).find();
        if (isNumber) {
            final long time = Long.parseLong(trimContent);
            if (type == Date.class) {
                field.set(obj, new Date(time));
                return true;
            }
            LocalDateTime localDateTime = LocalDateTime.ofInstant(Instant.ofEpochSecond(time), TimeZone
                    .getDefault().toZoneId());
            if (type == LocalDateTime.class) {
                field.set(obj, localDateTime);
                return true;
            }
            field.set(obj, localDateTime.toLocalDate());
        } else {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            String dateFormat = DEFAULT_DATE_FORMAT;
            if (Objects.nonNull(excelColumn) && StringUtil.isNotBlank(excelColumn.dateFormatPattern())) {
                dateFormat = excelColumn.dateFormatPattern();
            }
            if (type == Date.class) {
                SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
                field.set(obj, sdf.parse(trimContent));
                return true;
            }
            DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(dateFormat);
            if (type == LocalDateTime.class) {
                field.set(obj, LocalDateTime.parse(trimContent, dateTimeFormatter));
                return true;
            }
            field.set(obj, LocalDate.parse(trimContent, dateTimeFormatter));
        }
        return true;
    }
}
