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

import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.ConvertContext;
import com.github.liaochong.myexcel.core.converter.WriteConverter;
import com.github.liaochong.myexcel.utils.PropertyUtil;

import java.lang.reflect.Field;
import java.util.Properties;

/**
 * @author liaochong
 * @version 1.0
 */
public class MappingWriteConverter implements WriteConverter {

    private final WeakCache<String, Pair<Class, Object>> mappingCache = new WeakCache<>();

    @Override
    public boolean support(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        ExcelColumnMapping mapping = convertContext.excelColumnMappingMap.get(field);
        return mapping != null && !mapping.mapping.isEmpty();
    }

    @Override
    public Pair<Class, Object> convert(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        ExcelColumnMapping excelColumnMapping = convertContext.excelColumnMappingMap.get(field);
        String cacheKey = excelColumnMapping.mapping + "->" + fieldVal;
        Pair<Class, Object> mapping = mappingCache.get(cacheKey);
        if (mapping != null) {
            return mapping;
        }
        Properties properties = PropertyUtil.getProperties(excelColumnMapping);
        String property = properties.getProperty(fieldVal.toString());
        if (property == null) {
            return Pair.of(fieldType, fieldVal);
        }
        Pair<Class, Object> result = Pair.of(String.class, property);
        mappingCache.cache(cacheKey, result);
        return result;
    }
}
