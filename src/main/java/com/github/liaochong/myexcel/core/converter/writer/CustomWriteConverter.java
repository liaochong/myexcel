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
import com.github.liaochong.myexcel.core.converter.CustomConverter;
import com.github.liaochong.myexcel.core.converter.DefaultCustomConverter;
import com.github.liaochong.myexcel.core.converter.WriteConverter;

import java.lang.reflect.Field;

/**
 * 自定义映射关系
 *
 * @author liaochong
 * @version 1.0
 */
public class CustomWriteConverter implements WriteConverter {

    private WeakCache<Class, CustomConverter> cache = new WeakCache<>();

    @Override
    public boolean support(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        ExcelColumnMapping mapping = convertContext.getExcelColumnMappingMap().get(field);
        return mapping != null && mapping.getCustomConverter() != null && mapping.getCustomConverter() != DefaultCustomConverter.class;
    }

    @Override
    public Pair<Class, Object> convert(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        ExcelColumnMapping excelColumnMapping = convertContext.getExcelColumnMappingMap().get(field);
        Class<? extends CustomConverter> converter = excelColumnMapping.getCustomConverter();
        // 尝试绑定上下文中是否存在
        Object target = convertContext.getConfiguration().getApplicationBeans().get(converter);
        if (target != null) {
            Object result = ((CustomConverter) target).convert(fieldVal);
            return Pair.of(result.getClass(), result);
        }
        if (cache.get(converter) == null) {
            CustomConverter customConverter;
            try {
                customConverter = converter.newInstance();
            } catch (Exception e) {
                throw new IllegalStateException(e);
            }
            cache.cache(converter, customConverter);
        }
        Object result = cache.get(converter).convert(fieldVal);
        return Pair.of(result.getClass(), result);
    }
}
