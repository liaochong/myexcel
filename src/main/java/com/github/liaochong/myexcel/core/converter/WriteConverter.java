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

import java.lang.reflect.Field;

/**
 * @author liaochong
 * @version 1.0
 */
public interface WriteConverter {

    /**
     * 转换
     *
     * @param field    字段
     * @param fieldVal 字段对应的值
     * @return T
     */
    Pair<Class, Object> convert(Field field, Object fieldVal);

    /**
     * 是否支持转换
     *
     * @param field    字段
     * @param fieldVal 字段值
     * @return true/false
     */
    boolean support(Field field, Object fieldVal);

}
