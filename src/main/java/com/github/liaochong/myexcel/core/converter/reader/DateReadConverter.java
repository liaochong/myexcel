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

import com.github.liaochong.myexcel.core.cache.WeakCache;

import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;

/**
 * Date读取转换器
 *
 * @author liaochong
 * @version 1.0
 */
public class DateReadConverter extends AbstractReadConverter<Date> {

    private WeakCache<String, SimpleDateFormat> simpleDateFormatWeakCache = new WeakCache<>();

    @Override
    public Date doConvert(String v, Field field) {
        if (isNumber(v)) {
            final long time = Long.parseLong(v);
            return new Date(time);
        }
        String dateFormatPattern = getDateFormatPattern(field);
        SimpleDateFormat sdf = simpleDateFormatWeakCache.get(dateFormatPattern);
        if (Objects.isNull(sdf)) {
            sdf = new SimpleDateFormat(dateFormatPattern);
            simpleDateFormatWeakCache.cache(dateFormatPattern, sdf);
        }
        try {
            return sdf.parse(v);
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }
}
