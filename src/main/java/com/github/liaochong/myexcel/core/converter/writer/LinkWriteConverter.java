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
import com.github.liaochong.myexcel.core.constant.LinkEmail;
import com.github.liaochong.myexcel.core.constant.LinkUrl;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.WriteConverter;

import java.lang.reflect.Field;

/**
 * 超链接
 *
 * @author liaochong
 * @version 1.0
 */
public class LinkWriteConverter implements WriteConverter {

    @Override
    public Pair<Class, Object> convert(Field field, Object fieldVal) {
        String link = field.getAnnotation(ExcelColumn.class).link();
        String[] splits = link.split(":");
        if (splits.length == 1) {
            // 默认为url
            return Pair.of(LinkUrl.class, link);
        }
        switch (splits[0]) {
            case "url":
                return Pair.of(LinkUrl.class, splits[1]);
            case "email":
                return Pair.of(LinkEmail.class, splits[1]);
            default:
                throw new IllegalArgumentException("Illegal link type, only URL or email");
        }
    }

    @Override
    public boolean support(Field field, Object fieldVal) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        return excelColumn != null && !"".equals(excelColumn.link());
    }
}
