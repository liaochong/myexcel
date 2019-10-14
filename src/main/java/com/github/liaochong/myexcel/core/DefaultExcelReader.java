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

import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.converter.ReadConverterContext;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelReader<T> {

    private static final int DEFAULT_SHEET_INDEX = 0;

    private Class<T> dataType;

    private int sheetIndex = DEFAULT_SHEET_INDEX;

    private Predicate<Row> rowFilter = row -> true;

    private Predicate<T> beanFilter = bean -> true;

    private Workbook wb;

    private BiFunction<Throwable, ReadContext, Boolean> exceptionFunction = (e, c) -> false;

    private DefaultExcelReader(Class<T> dataType) {
        this.dataType = dataType;
    }

    public static <T> DefaultExcelReader<T> of(@NonNull Class<T> clazz) {
        return new DefaultExcelReader<>(clazz);
    }

    public DefaultExcelReader<T> sheet(int index) {
        if (index >= 0) {
            this.sheetIndex = index;
        } else {
            throw new IllegalArgumentException("Sheet index must be greater than or equal to 0");
        }
        return this;
    }

    public DefaultExcelReader<T> rowFilter(@NonNull Predicate<Row> rowFilter) {
        this.rowFilter = rowFilter;
        return this;
    }

    public DefaultExcelReader<T> beanFilter(@NonNull Predicate<T> beanFilter) {
        this.beanFilter = beanFilter;
        return this;
    }

    public DefaultExcelReader<T> exceptionally(BiFunction<Throwable, ReadContext, Boolean> exceptionFunction) {
        this.exceptionFunction = exceptionFunction;
        return this;
    }

    public List<T> read(@NonNull InputStream fileInputStream) throws Exception {
        return this.read(fileInputStream, null);
    }

    public List<T> read(@NonNull InputStream fileInputStream, String password) throws Exception {
        Map<Integer, Field> fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        if (fieldMap.isEmpty()) {
            return Collections.emptyList();
        }
        try {
            Sheet sheet = getSheetOfInputStream(fileInputStream, password);
            return getDataFromFile(sheet, fieldMap);
        } finally {
            if (Objects.nonNull(wb)) {
                wb.close();
            }
        }
    }

    public List<T> read(@NonNull File file) throws Exception {
        return this.read(file, null);
    }

    public List<T> read(@NonNull File file, String password) throws Exception {
        if (!file.getName().endsWith(Constants.XLSX) && !file.getName().endsWith(Constants.XLS)) {
            throw new IllegalArgumentException("Support only. xls and. xlsx suffix files");
        }
        Map<Integer, Field> fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        if (fieldMap.isEmpty()) {
            return Collections.emptyList();
        }
        try {
            Sheet sheet = getSheetOfFile(file, password);
            return getDataFromFile(sheet, fieldMap);
        } finally {
            if (Objects.nonNull(wb)) {
                wb.close();
            }
        }
    }

    public void readThen(@NonNull InputStream fileInputStream, Consumer<T> consumer) throws Exception {
        readThen(fileInputStream, null, consumer);
    }

    public void readThen(@NonNull InputStream fileInputStream, String password, Consumer<T> consumer) throws Exception {
        Map<Integer, Field> fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        if (fieldMap.isEmpty()) {
            return;
        }
        try {
            Sheet sheet = getSheetOfInputStream(fileInputStream, password);
            readThenConsume(sheet, fieldMap, consumer, null);
        } finally {
            if (Objects.nonNull(wb)) {
                wb.close();
            }
        }
    }

    public void readThen(@NonNull File file, Consumer<T> consumer) throws Exception {
        readThen(file, null, consumer);
    }

    public void readThen(@NonNull File file, String password, Consumer<T> consumer) throws Exception {
        if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls")) {
            throw new IllegalArgumentException("Support only. xls and. xlsx suffix files");
        }
        Map<Integer, Field> fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        if (fieldMap.isEmpty()) {
            return;
        }
        try {
            Sheet sheet = getSheetOfFile(file, password);
            readThenConsume(sheet, fieldMap, consumer, null);
        } finally {
            if (Objects.nonNull(wb)) {
                wb.close();
            }
        }
    }

    public void readThen(@NonNull InputStream fileInputStream, Function<T, Boolean> function) throws Exception {
        readThen(fileInputStream, null, function);
    }

    public void readThen(@NonNull InputStream fileInputStream, String password, Function<T, Boolean> function) throws Exception {
        Map<Integer, Field> fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        if (fieldMap.isEmpty()) {
            return;
        }
        try {
            Sheet sheet = getSheetOfInputStream(fileInputStream, password);
            readThenConsume(sheet, fieldMap, null, function);
        } finally {
            if (Objects.nonNull(wb)) {
                wb.close();
            }
        }
    }

    public void readThen(@NonNull File file, Function<T, Boolean> function) throws Exception {
        readThen(file, null, function);
    }

    public void readThen(@NonNull File file, String password, Function<T, Boolean> function) throws Exception {
        if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls")) {
            throw new IllegalArgumentException("Support only. xls and. xlsx suffix files");
        }
        Map<Integer, Field> fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        if (fieldMap.isEmpty()) {
            return;
        }
        try {
            Sheet sheet = getSheetOfFile(file, password);
            readThenConsume(sheet, fieldMap, null, function);
        } finally {
            if (Objects.nonNull(wb)) {
                wb.close();
            }
        }
    }

    private Sheet getSheetOfInputStream(@NonNull InputStream fileInputStream, String password) throws IOException {
        if (StringUtil.isBlank(password)) {
            wb = WorkbookFactory.create(fileInputStream);
        } else {
            wb = WorkbookFactory.create(fileInputStream, password);
        }
        return wb.getSheetAt(sheetIndex);
    }

    private Sheet getSheetOfFile(@NonNull File file, String password) throws IOException {
        if (StringUtil.isBlank(password)) {
            wb = WorkbookFactory.create(file);
        } else {
            wb = WorkbookFactory.create(file, password);
        }
        return wb.getSheetAt(sheetIndex);
    }

    private List<T> getDataFromFile(Sheet sheet, Map<Integer, Field> fieldMap) {
        long startTime = System.currentTimeMillis();
        final int firstRowNum = sheet.getFirstRowNum();
        final int lastRowNum = sheet.getLastRowNum();
        log.info("FirstRowNum:{},LastRowNum:{}", firstRowNum, lastRowNum);
        if (lastRowNum < 0) {
            log.info("Reading excel takes {} milliseconds", System.currentTimeMillis() - startTime);
            return Collections.emptyList();
        }
        DataFormatter formatter = new DataFormatter();
        List<T> result = new LinkedList<>();
        for (int i = firstRowNum; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (Objects.isNull(row)) {
                log.info("Row of {} is null,it will be ignored.", i);
                continue;
            }
            boolean noMatchResult = rowFilter.negate().test(row);
            if (noMatchResult) {
                log.info("Row of {} does not meet the filtering criteria, it will be ignored.", i);
                continue;
            }
            int lastColNum = row.getLastCellNum();
            if (lastColNum < 0) {
                continue;
            }
            T obj = instanceObj(fieldMap, formatter, row);
            if (beanFilter.test(obj)) {
                result.add(obj);
            }
        }
        log.info("Reading excel takes {} milliseconds", System.currentTimeMillis() - startTime);
        return result;
    }

    private void readThenConsume(Sheet sheet, Map<Integer, Field> fieldMap, Consumer<T> consumer, Function<T, Boolean> function) {
        long startTime = System.currentTimeMillis();
        final int firstRowNum = sheet.getFirstRowNum();
        final int lastRowNum = sheet.getLastRowNum();
        log.info("FirstRowNum:{},LastRowNum:{}", firstRowNum, lastRowNum);
        if (lastRowNum < 0) {
            log.info("Reading excel takes {} milliseconds", System.currentTimeMillis() - startTime);
            return;
        }
        DataFormatter formatter = new DataFormatter();
        for (int i = firstRowNum; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (Objects.isNull(row)) {
                log.info("Row of {} is null,it will be ignored.", i);
                continue;
            }
            boolean noMatchResult = rowFilter.negate().test(row);
            if (noMatchResult) {
                log.info("Row of {} does not meet the filtering criteria, it will be ignored.", i);
                continue;
            }
            int lastColNum = row.getLastCellNum();
            if (lastColNum < 0) {
                continue;
            }
            T obj = instanceObj(fieldMap, formatter, row);
            if (beanFilter.test(obj)) {
                if (Objects.nonNull(consumer)) {
                    consumer.accept(obj);
                } else if (Objects.nonNull(function)) {
                    Boolean noStop = function.apply(obj);
                    if (!noStop) {
                        break;
                    }
                }
            }
        }
        log.info("Reading excel takes {} milliseconds", System.currentTimeMillis() - startTime);
    }

    private T instanceObj(Map<Integer, Field> fieldMap, DataFormatter formatter, Row row) {
        T obj;
        try {
            obj = dataType.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
        fieldMap.forEach((key, field) -> {
            Cell cell = row.getCell(key, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (Objects.isNull(cell)) {
                return;
            }
            String content = formatter.formatCellValue(cell);
            ReadContext context = new ReadContext(field, content, row.getRowNum(), key);
            ReadConverterContext.convert(obj, context, exceptionFunction);
        });
        return obj;
    }
}
