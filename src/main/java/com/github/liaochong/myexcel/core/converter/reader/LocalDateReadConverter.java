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

import java.lang.reflect.Field;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.TimeZone;
import java.util.regex.Pattern;

/**
 * LocalDate读取转换器
 *
 * @author liaochong
 * @version 1.0
 */
public class LocalDateReadConverter extends AbstractReadConverter<LocalDate> {

    private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    private static final Pattern NUMBER_PATTERN = Pattern.compile("^\\d+$");

    @Override
    public LocalDate doConvert(String v, Field field) {
        boolean isNumber = NUMBER_PATTERN.matcher(v).find();
        if (isNumber) {
            final long time = Long.parseLong(v);

            LocalDateTime localDateTime = LocalDateTime.ofInstant(Instant.ofEpochSecond(time), TimeZone
                    .getDefault().toZoneId());
            return localDateTime.toLocalDate();
        }
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(getDateFormatPattern(field, DEFAULT_DATE_FORMAT));
        return LocalDate.parse(v, dateTimeFormatter);
    }
}
