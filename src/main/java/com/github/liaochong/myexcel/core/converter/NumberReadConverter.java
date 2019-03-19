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

import com.github.liaochong.myexcel.utils.StringUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;

/**
 * @author liaochong
 * @version 1.0
 */
public class NumberReadConverter implements ReadConverter {

    @Override
    public void convert(String content, Field field, Object obj) throws Exception {
        if (StringUtil.isBlank(content)) {
            return;
        }
        Class<?> type = field.getType();
        if (type != Double.class && type != double.class
                && type != Float.class && type != float.class
                && type != Long.class && type != long.class
                && type != Integer.class && type != int.class
                && type != Short.class && type != short.class
                && type != Byte.class && type != byte.class
                && type != BigDecimal.class) {
            return;
        }
        String trimContent = content.trim();
        String realValue = new BigDecimal(trimContent).toPlainString();

        field.setAccessible(true);
        if (type == Double.class || type == double.class) {
            field.set(obj, Double.parseDouble(realValue));
            return;
        }
        if (type == Float.class || type == float.class) {
            field.set(obj, Float.parseFloat(realValue));
            return;
        }
        if (type == Long.class || type == long.class) {
            field.set(obj, Long.parseLong(realValue));
            return;
        }
        if (type == Integer.class || type == int.class) {
            field.set(obj, Integer.parseInt(realValue));
            return;
        }
        if (type == Short.class || type == short.class) {
            field.set(obj, Short.parseShort(realValue));
            return;
        }
        if (type == Byte.class || type == byte.class) {
            field.set(obj, Byte.parseByte(realValue));
            return;
        }
        field.set(obj, new BigDecimal(realValue));
    }
}
