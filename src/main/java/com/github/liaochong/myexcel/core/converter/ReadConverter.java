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

import java.lang.reflect.Field;

/**
 * 转换接口
 *
 * @author liaochong
 * @version 1.0
 */
public interface ReadConverter<E, T> {

    /**
     * 转换
     *
     * @param obj            被转换对象
     * @param field          字段，提供额外信息
     * @param convertContext 转换上下文
     * @return 转换结果
     */
    T convert(E obj, Field field, ConvertContext convertContext);
}
