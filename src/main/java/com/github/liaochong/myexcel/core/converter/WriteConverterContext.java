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

import com.github.liaochong.myexcel.core.ConvertContext;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.constant.AllConverter;
import com.github.liaochong.myexcel.core.constant.CsvConverter;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.writer.BigDecimalWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.DateTimeWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.DropDownListWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.ImageWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.LinkWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.MappingWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.OriginalWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.StringWriteConverter;
import com.github.liaochong.myexcel.utils.ReflectUtil;

import javax.lang.model.type.NullType;
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

    private static final Pair<? extends Class, Object> NULL_PAIR = Pair.of(NullType.class, null);

    private static final List<Pair<Class, WriteConverter>> WRITE_CONVERTER_CONTAINER = new ArrayList<>();

    private static final WeakCache<Field, WriteConverter> CONVERTER_CACHE = new WeakCache<>();

    private static final OriginalWriteConverter ORIGINAL_WRITE_CONVERTER = new OriginalWriteConverter();

    static {
        WRITE_CONVERTER_CONTAINER.add(Pair.of(CsvConverter.class, new DateTimeWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new StringWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new BigDecimalWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new DropDownListWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new LinkWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new MappingWriteConverter()));
        WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, new ImageWriteConverter()));
    }

    public static synchronized void registering(WriteConverter... writeConverters) {
        Objects.requireNonNull(writeConverters);
        for (WriteConverter writeConverter : writeConverters) {
            WRITE_CONVERTER_CONTAINER.add(Pair.of(AllConverter.class, writeConverter));
        }
    }

    public static Pair<? extends Class, Object> convert(Field field, Object object, Class converterType, ConvertContext convertContext) {
        Object result = ReflectUtil.getFieldValue(object, field);
        if (result == null) {
            return NULL_PAIR;
        }
        WriteConverter writeConverter = CONVERTER_CACHE.get(field);
        if (writeConverter == null) {
            Optional<WriteConverter> writeConverterOptional = WRITE_CONVERTER_CONTAINER.stream()
                    .filter(pair -> (pair.getKey() == converterType || pair.getKey() == AllConverter.class) && pair.getValue().support(field, result, convertContext))
                    .map(Pair::getValue)
                    .findFirst();
            writeConverter = writeConverterOptional.isPresent() ? writeConverterOptional.get() : ORIGINAL_WRITE_CONVERTER;
            CONVERTER_CACHE.cache(field, writeConverter);
        }
        return writeConverter.convert(field, result, convertContext);
    }
}
