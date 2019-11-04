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
import com.github.liaochong.myexcel.exception.StopReadException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * sax处理
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
class XSSFSaxReadHandler<T> extends AbstractReadHandler<T> implements XSSFSheetXMLHandler.SheetContentsHandler {

    private final Map<Integer, Field> fieldMap;

    private List<T> result;

    private T obj;

    private Class<T> dataType;

    private Consumer<T> consumer;

    private Function<T, Boolean> function;

    private Predicate<Row> rowFilter;

    private Predicate<T> beanFilter;

    private Row currentRow;

    private int count;

    private BiFunction<Throwable, ReadContext, Boolean> exceptionFunction;

    public XSSFSaxReadHandler(
            Map<Integer, Field> fieldMap,
            List<T> result,
            SaxExcelReader.ReadConfig<T> readConfig) {
        this.fieldMap = fieldMap;
        this.result = result;
        this.dataType = readConfig.getDataType();
        this.consumer = readConfig.getConsumer();
        this.function = readConfig.getFunction();
        this.rowFilter = readConfig.getRowFilter();
        this.beanFilter = readConfig.getBeanFilter();
        this.exceptionFunction = readConfig.getExceptionFunction();
    }

    @Override
    public void startRow(int rowNum) {
        currentRow = new Row(rowNum);
        obj = this.newInstance(dataType);
    }

    @Override
    public void endRow(int rowNum) {
        if (!rowFilter.test(currentRow)) {
            return;
        }
        if (!beanFilter.test(obj)) {
            return;
        }
        count++;
        if (consumer != null) {
            consumer.accept(obj);
        } else if (function != null) {
            Boolean noStop = function.apply(obj);
            if (!noStop) {
                throw new StopReadException();
            }
        } else {
            result.add(obj);
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    public void cell(String cellReference, String formattedValue,
                     XSSFComment comment) {
        if (!rowFilter.test(currentRow)) {
            return;
        }
        if (cellReference == null) {
            return;
        }
        int thisCol = (new CellReference(cellReference)).getCol();
        if (isMapType) {
            ((Map<Integer, String>) obj).put(thisCol, formattedValue);
            return;
        }
        Field field = fieldMap.get(thisCol);
        if (field == null) {
            return;
        }
        ReadContext<T> context = new ReadContext<>(obj, field, formattedValue, currentRow.getRowNum(), thisCol);
        ReadConverterContext.convert(obj, context, exceptionFunction);
    }

    @Override
    public void endSheet() {
        log.info("Import completed, total number of rows {}", count);
    }
}
