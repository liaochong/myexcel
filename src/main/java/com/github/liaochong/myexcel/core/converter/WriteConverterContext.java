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

import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.writer.BigDecimalWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.DateTimeWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.DropDownListWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.LinkWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.MappingWriteConverter;
import com.github.liaochong.myexcel.core.converter.writer.StringWriteConverter;
import com.github.liaochong.myexcel.utils.ReflectUtil;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * @author liaochong
 * @version 1.0
 */
public class WriteConverterContext {

    private static final Pair<? extends Class, Object> NULL_PAIR = Pair.of(String.class, null);

    private static final List<WriteConverter> WRITE_CONVERTER_CONTAINER = new ArrayList<>();

    static {
        WRITE_CONVERTER_CONTAINER.add(new DateTimeWriteConverter());
        WRITE_CONVERTER_CONTAINER.add(new StringWriteConverter());
        WRITE_CONVERTER_CONTAINER.add(new BigDecimalWriteConverter());
        WRITE_CONVERTER_CONTAINER.add(new DropDownListWriteConverter());
        WRITE_CONVERTER_CONTAINER.add(new LinkWriteConverter());
        WRITE_CONVERTER_CONTAINER.add(new MappingWriteConverter());
    }

    public static synchronized void registering(WriteConverter... writeConverters) {
        Objects.requireNonNull(writeConverters);
        Collections.addAll(WRITE_CONVERTER_CONTAINER, writeConverters);
    }

    public static Pair<? extends Class, Object> convert(Field field, Object object) {
        Object result = ReflectUtil.getFieldValue(object, field);
        if (result == null) {
            return NULL_PAIR;
        }
        Optional<WriteConverter> writeConverterOptional = WRITE_CONVERTER_CONTAINER.stream()
                .filter(writeConverter -> writeConverter.support(field, result))
                .findFirst();
        return writeConverterOptional.isPresent() ? writeConverterOptional.get().convert(field, result) : Pair.of(field.getType(), result);
    }
}
