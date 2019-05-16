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
package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.converter.ReadConverterContext;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Consumer;

/**
 * sax处理
 *
 * @author liaochong
 * @version 1.0
 */
class SaxHandler<T> implements XSSFSheetXMLHandler.SheetContentsHandler {

    private final Map<Integer, Field> fieldMap;

    private List<T> result;

    private T obj;

    private Class<T> dataType;

    private Consumer<T> consumer;

    public SaxHandler(Class<T> dataType, Map<Integer, Field> fieldMap, List<T> result, Consumer<T> consumer) {
        this.fieldMap = fieldMap;
        this.result = result;
        this.dataType = dataType;
        this.consumer = consumer;
    }

    @Override
    public void startRow(int rowNum) {
        try {
            obj = dataType.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void endRow(int rowNum) {
        if (Objects.isNull(consumer)) {
            consumer.accept(obj);
        } else {
            result.add(obj);
        }
    }

    @Override
    public void cell(String cellReference, String formattedValue,
                     XSSFComment comment) {
        if (cellReference == null) {
            return;
        }
        int thisCol = (new CellReference(cellReference)).getCol();
        Field field = fieldMap.get(thisCol);
        if (Objects.isNull(field)) {
            throw new RuntimeException();
        }
        ReadConverterContext.convert(formattedValue, field, obj);
    }
}
