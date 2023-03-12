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

import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.annotation.MultiColumn;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.context.Hyperlink;
import com.github.liaochong.myexcel.core.context.ReadContext;
import com.github.liaochong.myexcel.core.converter.reader.BigDecimalReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.BoolReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.DateReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.HyperlinkReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.LocalDateReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.LocalDateTimeReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.LocalTimeReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.NumberReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.StringReadConverter;
import com.github.liaochong.myexcel.core.converter.reader.TimestampReadConverter;
import com.github.liaochong.myexcel.exception.ExcelReadException;
import com.github.liaochong.myexcel.exception.SaxReadException;
import com.github.liaochong.myexcel.utils.PropertyUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
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

    private static final Map<Class<?>, ReadConverter<?>> READ_CONVERTERS = new HashMap<>();

    private static final WeakCache<Field, Properties> MAPPING_CACHE = new WeakCache<>();

    private static final Properties EMPTY_PROPERTIES = new Properties();

    static {
        READ_CONVERTERS.put(Hyperlink.class, new HyperlinkReadConverter());

        BoolReadConverter boolReadConverter = new BoolReadConverter();
        READ_CONVERTERS.put(Boolean.class, boolReadConverter);
        READ_CONVERTERS.put(boolean.class, boolReadConverter);

        READ_CONVERTERS.put(Date.class, new DateReadConverter());
        READ_CONVERTERS.put(LocalDate.class, new LocalDateReadConverter());
        READ_CONVERTERS.put(LocalDateTime.class, new LocalDateTimeReadConverter());
        READ_CONVERTERS.put(LocalTime.class, new LocalTimeReadConverter());

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

    public static boolean support(Class<?> clazz) {
        return READ_CONVERTERS.get(clazz) != null;
    }

    public synchronized ReadConverterContext registering(Class<?> clazz, ReadConverter<?> readConverter) {
        READ_CONVERTERS.putIfAbsent(clazz, readConverter);
        return this;
    }

    @SuppressWarnings("unchecked")
    public static void convert(Object obj, ReadContext<?> readContext, BiFunction<Throwable, ReadContext, Boolean> exceptionFunction) {
        ReadConverter<?> readConverter = READ_CONVERTERS.get(readContext.getField().getType());
        if (readConverter == null) {
            MultiColumn multiColumn = readContext.getField().getAnnotation(MultiColumn.class);
            if (multiColumn != null) {
                readConverter = READ_CONVERTERS.get(multiColumn.classType());
            }
            if (readConverter == null) {
                throw new IllegalStateException("No suitable type converter was found,Field=" + readContext.getField().getName() + ".");
            }
        }
        Object value = null;
        try {
            Properties properties = MAPPING_CACHE.get(readContext.getField());
            if (properties == null) {
                ExcelColumnMapping mapping = readContext.getConvertContext().excelColumnMappingMap.get(readContext.getField());
                if (mapping != null && !mapping.mapping.isEmpty()) {
                    properties = PropertyUtil.getReverseProperties(mapping);
                } else {
                    properties = EMPTY_PROPERTIES;
                }
                MAPPING_CACHE.cache(readContext.getField(), properties);
            }
            String mappingVal = properties.getProperty(readContext.getVal());
            if (mappingVal != null) {
                readContext.setVal(mappingVal);
            }
            value = readConverter.convert(readContext);
        } catch (Exception e) {
            Boolean toContinue = exceptionFunction.apply(e, readContext);
            if (!toContinue) {
                throw new ExcelReadException("Failed to convert content,field:[" + readContext.getField().getDeclaringClass().getName() + "#" + readContext.getField().getName() + "],content:[" + readContext.getVal() + "],rowNum:[" + readContext.getRowNum() + "]", e);
            }
        }
        if (value == null) {
            return;
        }
        try {
            if (obj instanceof List) {
                ((List) obj).add(value);
            } else {
                readContext.getField().set(obj, value);
            }
        } catch (Exception e) {
            throw new SaxReadException("Failed to set the " + readContext.getField().getDeclaringClass().getName() + "#" + readContext.getField().getName() + " field value to " + readContext.getVal(), e);
        }
    }
}
