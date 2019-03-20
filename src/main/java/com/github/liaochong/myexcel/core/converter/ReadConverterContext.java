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
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public class ReadConverterContext {

    private static List<ReadConverter> readConverterList = new ArrayList<>();

    static {
        readConverterList.add(new NumberReadConverter());
        readConverterList.add(new DateReadConverter());
        readConverterList.add(new BoolReadConverter());
        readConverterList.add(new StringReadConverter());
    }

    public static ReadConverterContext getInstance() {
        return new ReadConverterContext();
    }

    public synchronized ReadConverterContext registering(ReadConverter... readConverters) {
        Objects.requireNonNull(readConverters);
        Collections.addAll(readConverterList, readConverters);
        return this;
    }

    public void convert(String content, Field field, Object obj) {
        readConverterList.forEach(readConverter -> {
            try {
                readConverter.convert(content, field, obj);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        });
    }
}
