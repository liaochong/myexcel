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

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.WriteConverter;
import com.github.liaochong.myexcel.utils.PropertyUtil;

import java.lang.reflect.Field;
import java.util.Properties;

/**
 * @author liaochong
 * @version 1.0
 */
public class MappingWriteConverter implements WriteConverter {

    private WeakCache<String, Pair<Class, Object>> mappingCache = new WeakCache<>();

    @Override
    public boolean support(Field field, Object fieldVal) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        return excelColumn != null && !excelColumn.mapping().isEmpty();
    }

    @Override
    public Pair<Class, Object> convert(Field field, Object fieldVal) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        String cacheKey = excelColumn.mapping() + "->" + fieldVal;
        Pair<Class, Object> mapping = mappingCache.get(cacheKey);
        if (mapping != null) {
            return mapping;
        }
        Properties properties = PropertyUtil.getProperties(excelColumn);
        String property = properties.getProperty(fieldVal.toString());
        if (property == null) {
            return Pair.of(field.getType(), fieldVal);
        }
        Pair<Class, Object> result = Pair.of(String.class, property);
        mappingCache.cache(cacheKey, result);
        return result;
    }
}
