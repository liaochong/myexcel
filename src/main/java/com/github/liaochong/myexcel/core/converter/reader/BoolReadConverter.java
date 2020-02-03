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
package com.github.liaochong.myexcel.core.converter.reader;

import com.github.liaochong.myexcel.core.ConvertContext;
import com.github.liaochong.myexcel.core.constant.Constants;

import java.lang.reflect.Field;
import java.util.Objects;

/**
 * 布尔转换器
 *
 * @author liaochong
 * @version 1.0
 */
public class BoolReadConverter extends AbstractReadConverter<Boolean> {

    @Override
    public Boolean doConvert(String v, Field field, ConvertContext convertContext) {
        if (Objects.equals(Constants.ONE, v) || v.equalsIgnoreCase(Constants.TRUE)) {
            return Boolean.TRUE;
        }
        if (Objects.equals(Constants.ZERO, v) || v.equalsIgnoreCase(Constants.FALSE)) {
            return Boolean.FALSE;
        }
        throw new IllegalStateException("Cell content does not match the type of field to be injected,field is " + field.getName() + ",value is \"" + v + "\"");
    }
}
