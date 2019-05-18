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

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * 读取转换器上下文
 *
 * @author liaochong
 * @version 1.0
 */
public class ReadConverterContext {

    private static final Map<Class<?>, Converter<String, ?>> READ_CONVERTERS = new HashMap<>();

    static {
        BoolReadConverter boolReadConverter = new BoolReadConverter();
        READ_CONVERTERS.put(Boolean.class, boolReadConverter);
        READ_CONVERTERS.put(boolean.class, boolReadConverter);

        READ_CONVERTERS.put(Date.class, new DateReadConverter());
        READ_CONVERTERS.put(LocalDate.class, new LocalDateReadConverter());
        READ_CONVERTERS.put(LocalDateTime.class, new LocalDateTimeReadConverter());

        DoubleReadConverter doubleReadConverter = new DoubleReadConverter();
        READ_CONVERTERS.put(Double.class, doubleReadConverter);
        READ_CONVERTERS.put(double.class, doubleReadConverter);

        FloatReadConverter floatReadConverter = new FloatReadConverter();
        READ_CONVERTERS.put(Float.class, floatReadConverter);
        READ_CONVERTERS.put(float.class, floatReadConverter);

        LongReadConverter longReadConverter = new LongReadConverter();
        READ_CONVERTERS.put(Long.class, longReadConverter);
        READ_CONVERTERS.put(long.class, longReadConverter);

        IntegerReadConverter integerReadConverter = new IntegerReadConverter();
        READ_CONVERTERS.put(Integer.class, integerReadConverter);
        READ_CONVERTERS.put(int.class, integerReadConverter);

        ShortReadConverter shortReadConverter = new ShortReadConverter();
        READ_CONVERTERS.put(Short.class, shortReadConverter);
        READ_CONVERTERS.put(short.class, shortReadConverter);

        ByteReadConverter byteReadConverter = new ByteReadConverter();
        READ_CONVERTERS.put(Byte.class, byteReadConverter);
        READ_CONVERTERS.put(byte.class, byteReadConverter);

        READ_CONVERTERS.put(BigDecimal.class, new BigDecimalReadConverter());
        READ_CONVERTERS.put(String.class, new StringReadConverter());
    }

    public synchronized ReadConverterContext registering(Class<?> clazz, Converter<String, ?> converter) {
        READ_CONVERTERS.putIfAbsent(clazz, converter);
        return this;
    }

    public static void convert(String content, Field field, Object obj) {
        Converter<String, ?> converter = READ_CONVERTERS.get(field.getType());
        if (Objects.isNull(converter)) {
            throw new IllegalStateException("No suitable type converter was found.");
        }
        Object value = converter.convert(content, field);
        if (Objects.isNull(value)) {
            return;
        }
        try {
            field.set(obj, value);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }
}
