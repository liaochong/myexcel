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

import com.github.liaochong.myexcel.core.constant.DropDownList;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.WriteConverter;

import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 下拉列表转换器
 *
 * @author liaochong
 * @version 1.0
 */
public class DropDownListWriteConverter implements WriteConverter {

    @Override
    public Pair<Class, Object> convert(Field field, Object fieldVal) {
        String content;
        if (field.getType() == List.class) {
            content = ((List<?>) fieldVal).stream().map(Object::toString).collect(Collectors.joining(","));
        } else {
            content = Stream.of(((Array) fieldVal)).map(Object::toString).collect(Collectors.joining(","));
        }
        return Pair.of(DropDownList.class, content);
    }

    @Override
    public boolean support(Field field, Object fieldVal) {
        return field.getType() == Array.class || field.getType() == List.class;
    }
}
