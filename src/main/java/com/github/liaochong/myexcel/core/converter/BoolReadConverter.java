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
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public class BoolReadConverter implements ReadConverter {

    @Override
    public void convert(String content, Field field, Object obj) throws Exception {
        if (StringUtil.isBlank(content)) {
            return;
        }
        if (field.getType() != Boolean.class && field.getType() != boolean.class) {
            return;
        }
        field.setAccessible(true);
        String trimContent = content.trim();
        if (Objects.equals("1", trimContent) || Objects.equals("true", trimContent)) {
            field.set(obj, true);
            return;
        }
        if (Objects.equals("0", trimContent) || Objects.equals("false", trimContent)) {
            field.set(obj, false);
            return;
        }
        throw new IllegalArgumentException();
    }
}
