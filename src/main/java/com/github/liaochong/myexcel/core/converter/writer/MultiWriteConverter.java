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
package com.github.liaochong.myexcel.core.converter.writer;

import com.github.liaochong.myexcel.core.ConvertContext;
import com.github.liaochong.myexcel.core.annotation.MultiColumn;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.WriteConverter;
import com.github.liaochong.myexcel.core.converter.WriteConverterContext;

import java.lang.reflect.Field;
import java.util.LinkedList;
import java.util.List;

/**
 * 聚合列转换器
 *
 * @author liaochong
 * @version 1.0
 */
public class MultiWriteConverter implements WriteConverter {

    public final List<Pair<Class, WriteConverter>> writeConverterContainer;

    public MultiWriteConverter(List<Pair<Class, WriteConverter>> writeConverterContainer) {
        this.writeConverterContainer = new LinkedList<>(writeConverterContainer);
    }

    @Override
    public Pair<Class, Object> convert(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        MultiColumn multiColumn = field.getAnnotation(MultiColumn.class);
        List<Object> result = new LinkedList<>();
        for (Object o : ((List) fieldVal)) {
            if (o == null) {
                result.add(null);
                continue;
            }
            WriteConverter writeConverter = WriteConverterContext.getWriteConverter(field, multiColumn.classType(), o, convertContext, writeConverterContainer);
            Pair<Class, Object> convertResult = writeConverter.convert(field, multiColumn.classType(), o, convertContext);
            result.add(convertResult.getValue());
        }
        return result.isEmpty() ? WriteConverterContext.NULL_PAIR : Pair.of(multiColumn.classType(), result);
    }

    @Override
    public boolean support(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        return field.isAnnotationPresent(MultiColumn.class) && field.getType() == List.class;
    }
}
