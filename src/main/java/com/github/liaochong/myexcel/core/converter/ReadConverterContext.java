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

import com.github.liaochong.myexcel.core.ConvertContext;
import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.ReadContext;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.converter.reader.BigDecimalReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.BoolReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.DateReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.LocalDateReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.LocalDateTimeReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.NumberReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.StringReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.TimestampReadConverter;
import com.github.liaochong.myexcel.exception.ExcelReadException;
import com.github.liaochong.myexcel.exception.SaxReadException;
import com.github.liaochong.myexcel.utils.PropertyUtil;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import java.util.function.BiFunction;

/**
 * 读取转换器上下文
 *
 * @author liaochong
 * @version 1.0
 */
public class ReadConverterContext {

    private static final Map<Class<?>, Converter<String, ?>> READ_CONVERTERS = new HashMap<>();

    private static final WeakCache<Field, Properties> MAPPING_CACHE = new WeakCache<>();

    private static final Properties EMPTY_PROPERTIES = new Properties();
    private static final Logger log = org.slf4j.LoggerFactory.getLogger(ReadConverterContext.class);

    static {
        BoolReadConverter boolReadConverter = new BoolReadConverter();
        READ_CONVERTERS.put(Boolean.class, boolReadConverter);
        READ_CONVERTERS.put(boolean.class, boolReadConverter);

        READ_CONVERTERS.put(Date.class, new DateReadConverter());
        READ_CONVERTERS.put(LocalDate.class, new LocalDateReadConverter());
        READ_CONVERTERS.put(LocalDateTime.class, new LocalDateTimeReadConverter());

        NumberReadConverter<Double> doubleReadConverter = NumberReadConverter.of(Double::valueOf);
        READ_CONVERTERS.put(Double.class, doubleReadConverter);
        READ_CONVERTERS.put(double.class, doubleReadConverter);

        NumberReadConverter<Float> floatReadConverter = NumberReadConverter.of(Float::valueOf);
        READ_CONVERTERS.put(Float.class, floatReadConverter);
        READ_CONVERTERS.put(float.class, floatReadConverter);

        NumberReadConverter<Long> longReadConverter = NumberReadConverter.of(Long::valueOf, true);
        READ_CONVERTERS.put(Long.class, longReadConverter);
        READ_CONVERTERS.put(long.class, longReadConverter);

        NumberReadConverter<Integer> integerReadConverter = NumberReadConverter.of(Integer::valueOf, true);
        READ_CONVERTERS.put(Integer.class, integerReadConverter);
        READ_CONVERTERS.put(int.class, integerReadConverter);

        NumberReadConverter<Short> shortReadConverter = NumberReadConverter.of(Short::valueOf, true);
        READ_CONVERTERS.put(Short.class, shortReadConverter);
        READ_CONVERTERS.put(short.class, shortReadConverter);

        NumberReadConverter<Byte> byteReadConverter = NumberReadConverter.of(Byte::valueOf, true);
        READ_CONVERTERS.put(Byte.class, byteReadConverter);
        READ_CONVERTERS.put(byte.class, byteReadConverter);

        READ_CONVERTERS.put(BigDecimal.class, new BigDecimalReadConverter());
        READ_CONVERTERS.put(String.class, new StringReadConverter());

        READ_CONVERTERS.put(Timestamp.class, new TimestampReadConverter());

        NumberReadConverter<BigInteger> bigIntegerReadConverter = NumberReadConverter.of(BigInteger::new, true);
        READ_CONVERTERS.put(BigInteger.class, bigIntegerReadConverter);
    }

    public synchronized ReadConverterContext registering(Class<?> clazz, Converter<String, ?> converter) {
        READ_CONVERTERS.putIfAbsent(clazz, converter);
        return this;
    }

    public static void convert(Object obj, ReadContext context, ConvertContext convertContext, BiFunction<Throwable, ReadContext, Boolean> exceptionFunction) {
        Converter<String, ?> converter = READ_CONVERTERS.get(context.getField().getType());
        if (converter == null) {
            throw new IllegalStateException("No suitable type converter was found.");
        }
        Object value = null;
        try {
            Properties properties = MAPPING_CACHE.get(context.getField());
            if (properties == null) {
                ExcelColumnMapping mapping = convertContext.getExcelColumnMappingMap().get(context.getField());
                if (mapping != null && !mapping.getMapping().isEmpty()) {
                    properties = PropertyUtil.getReverseProperties(mapping);
                } else {
                    properties = EMPTY_PROPERTIES;
                }
                MAPPING_CACHE.cache(context.getField(), properties);
            }
            String mappingVal = properties.getProperty(context.getVal());
            if (mappingVal != null) {
                context.setVal(mappingVal);
            }
            value = converter.convert(context.getVal(), context.getField(), convertContext);
        } catch (Exception e) {
            Boolean toContinue = exceptionFunction.apply(e, context);
            if (!toContinue) {
                throw new ExcelReadException("Failed to convert content,field:[" + context.getField().getDeclaringClass().getName() + "#" + context.getField().getName() + "],content:[" + context.getVal() + "],rowNum:[" + context.getRowNum() + "]", e);
            }
        }
        if (value == null) {
            return;
        }
        try {
            context.getField().set(obj, value);
        } catch (IllegalAccessException e) {
            throw new SaxReadException("Failed to set the " + context.getField().getDeclaringClass().getName() + "#" + context.getField().getName() + " field value to " + context.getVal(), e);
        }
    }
}
