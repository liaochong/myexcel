/*
 * Copyright 2017 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core.converter;

import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.constant.AllConverter;
import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.constant.CsvConverter;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.writer.BigDecimalWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.CustomWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.DateTimeWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.DropDownListWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.ImageWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.LinkWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.LocalTimeWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.MappingWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.MultiWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.OriginalWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.StringWriteConverter;
import com.github.liaochong.myexcel.utils.FieldDefinition;
import com.github.liaochong.myexcel.utils.ReflectUtil;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * @author liaochong
 * @version 1.0
 */
public class WriteConverterContext {

    private static final List<Pair<Class, WriteConverter>> WRITE_CONVERTER_CONTAINER = new ArrayList<>();

    private static final WeakCache<Pair<Field, Class<?>>, WriteConverter> EXCEL_CONVERTER_CACHE = new WeakCache<>();

    private static final WeakCache<Pair<Field, Class<?>>, WriteConverter> CSV_CONVERTER_CACHE = new WeakCache<>();

    private static final OriginalWriteConverter ORIGINAL_WRITE_CONVERTER = new OriginalWriteConverter();

    static {
        WRITE_CONVERTER_CONTAINER.add(Pair.of(CsvConverter.class, new DateTimeWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new LocalTimeWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new StringWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new BigDecimalWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new DropDownListWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new LinkWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new CustomWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new MappingWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new ImageWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new MultiWriteConverter(WRITE_CONVERTER_CONTAINER)));
    }

    public static synchronized void registering(WriteConverter... writeConverters) {
        Objects.requireNonNull(writeConverters);
        for (WriteConverter writeConverter : writeConverters) {
            WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, writeConverter));
        }
    }

    public static Pair<? extends Class, Object> convert(FieldDefinition fieldDefinition, Object object, ConvertContext convertContext) {
        Object result = getFieldVal(fieldDefinition, object);
        if (result == null) {
            return Constants.NULL_PAIR;
        }
        WriteConverter writeConverter = getWriteConverter(fieldDefinition.getField(), fieldDefinition.getField().getType(), result, convertContext, WRITE_CONVERTER_CONTAINER);
        return writeConverter.convert(fieldDefinition.getField(), fieldDefinition.getField().getType(), result, convertContext);
    }

    private static Object getFieldVal(FieldDefinition fieldDefinition, Object object) {
        Object result;
        if (fieldDefinition.getParentFields() == null
                || fieldDefinition.getParentFields().isEmpty() || fieldDefinition.getField().getType() == List.class) {
            result = ReflectUtil.getFieldValue(object, fieldDefinition);
        } else {
            Object prevObj;
            try {
                prevObj = object;
                for (int i = 0, size = fieldDefinition.getParentFields().size(); i < size; i++) {
                    Field parentField = fieldDefinition.getParentFields().get(i);
                    prevObj = parentField.get(prevObj);
                }
                if (prevObj == null) {
                    return null;
                }
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
            result = ReflectUtil.getFieldValue(prevObj, fieldDefinition);
        }
        return result;
    }

    public static WriteConverter getWriteConverter(Field field, Class<?> fieldType, Object result, ConvertContext convertContext, List<Pair<Class, WriteConverter>> writeConverterContainer) {
        WriteConverter writeConverter = convertContext.isConvertCsv ? CSV_CONVERTER_CACHE.get(Pair.of(field, fieldType)) : EXCEL_CONVERTER_CACHE.get(Pair.of(field, fieldType));
        if (writeConverter != null) {
            return writeConverter;
        }
        Optional<WriteConverter> writeConverterOptional = writeConverterContainer.stream()
                .filter(pair -> (pair.getKey() == convertContext.converterType || pair.getKey() == AllConverter.class) && pair.getValue().support(field, fieldType, result, convertContext))
                .map(Pair::getValue)
                .findFirst();
        writeConverter = writeConverterOptional.orElse(ORIGINAL_WRITE_CONVERTER);
        if (convertContext.isConvertCsv) {
            CSV_CONVERTER_CACHE.cache(Pair.of(field, fieldType), writeConverter);
        } else {
            EXCEL_CONVERTER_CACHE.cache(Pair.of(field, fieldType), writeConverter);
        }
        return writeConverter;
    }
}
