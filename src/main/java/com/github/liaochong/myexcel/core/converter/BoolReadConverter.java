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
    public boolean convert(String content, Field field, Object obj) throws Exception {
        if (StringUtil.isBlank(content)) {
            return false;
        }
        if (field.getType() != Boolean.class && field.getType() != boolean.class) {
            return false;
        }
        String trimContent = content.trim();
        if (Objects.equals("1", trimContent) || trimContent.equalsIgnoreCase("true")) {
            field.set(obj, true);
            return true;
        }
        if (Objects.equals("0", trimContent) || trimContent.equalsIgnoreCase("false")) {
            field.set(obj, false);
            return true;
        }
        throw new IllegalStateException("Cell content does not match the type of field to be injected,field is " + field.getName());
    }
}
