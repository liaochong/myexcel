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

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.converter.ReadConverterContext;
import com.github.liaochong.myexcel.core.parallel.ParallelContainer;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
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
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelReader {

    private static final int DEFAULT_SHEET_INDEX = 0;

    private Class<?> dataType;

    private int sheetIndex = DEFAULT_SHEET_INDEX;

    private Predicate<Row> rowFilter = row -> true;

    private boolean parallelRead;

    private DefaultExcelReader(Class<?> dataType) {
        this.dataType = dataType;
    }

    public static <T> DefaultExcelReader of(@NonNull Class<T> clazz) {
        return new DefaultExcelReader(clazz);
    }

    public DefaultExcelReader sheet(int index) {
        if (index >= 0) {
            this.sheetIndex = index;
        } else {
            throw new IllegalArgumentException("Sheet index must be greater than or equal to 0");
        }
        return this;
    }

    public DefaultExcelReader rowFilter(@NonNull Predicate<Row> rowFilter) {
        this.rowFilter = rowFilter;
        return this;
    }

    public DefaultExcelReader parallelRead() {
        this.parallelRead = true;
        return this;
    }

    public <T> List<T> read(@NonNull InputStream fileInputStream) throws Exception {
        return this.read(fileInputStream, null);
    }

    public <T> List<T> read(@NonNull InputStream fileInputStream, String password) throws Exception {
        Map<Integer, Field> fieldMap = getFieldMap();
        if (fieldMap.isEmpty()) {
            return Collections.emptyList();
        }
        Sheet sheet = getSheetOfInputStream(fileInputStream, password);
        return getDataFromFile(sheet, fieldMap);
    }

    public <T> List<T> read(@NonNull File file) throws Exception {
        return this.read(file, null);
    }

    public <T> List<T> read(@NonNull File file, String password) throws Exception {
        if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls")) {
            throw new IllegalArgumentException("Support only. xls and. xlsx suffix files");
        }
        Map<Integer, Field> fieldMap = getFieldMap();
        if (fieldMap.isEmpty()) {
            return Collections.emptyList();
        }
        Sheet sheet = getSheetOfFile(file, password);
        return getDataFromFile(sheet, fieldMap);
    }

    public <T> void readThen(@NonNull InputStream fileInputStream, Consumer<T> consumer) throws Exception {
        readThen(fileInputStream, null, consumer);
    }

    public <T> void readThen(@NonNull InputStream fileInputStream, String password, Consumer<T> consumer) throws Exception {
        Map<Integer, Field> fieldMap = getFieldMap();
        if (fieldMap.isEmpty()) {
            return;
        }
        Sheet sheet = getSheetOfInputStream(fileInputStream, password);
        readThenConsume(sheet, fieldMap, consumer);
    }


    public <T> void readThen(@NonNull File file, Consumer<T> consumer) throws Exception {
        readThen(file, null, consumer);
    }

    public <T> void readThen(@NonNull File file, String password, Consumer<T> consumer) throws Exception {
        if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls")) {
            throw new IllegalArgumentException("Support only. xls and. xlsx suffix files");
        }
        Map<Integer, Field> fieldMap = getFieldMap();
        if (fieldMap.isEmpty()) {
            return;
        }
        Sheet sheet = getSheetOfFile(file, password);
        readThenConsume(sheet, fieldMap, consumer);
    }

    private Sheet getSheetOfInputStream(@NonNull InputStream fileInputStream, String password) throws IOException {
        Workbook wb;
        if (StringUtil.isBlank(password)) {
            wb = WorkbookFactory.create(fileInputStream);
        } else {
            wb = WorkbookFactory.create(fileInputStream, password);
        }
        return wb.getSheetAt(sheetIndex);
    }

    private Sheet getSheetOfFile(@NonNull File file, String password) throws IOException {
        Workbook wb;
        if (StringUtil.isBlank(password)) {
            wb = WorkbookFactory.create(file);
        } else {
            wb = WorkbookFactory.create(file, password);
        }
        return wb.getSheetAt(sheetIndex);
    }

    private Map<Integer, Field> getFieldMap() {
        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
        List<Field> fields = classFieldContainer.getFieldsByAnnotation(ExcelColumn.class);
        if (fields.isEmpty()) {
            throw new IllegalStateException("There is no field with @ExcelColumn");
        }
        Map<Integer, Field> fieldMap = new HashMap<>(fields.size());
        for (Field field : fields) {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            int index = excelColumn.index();
            if (index < 0) {
                continue;
            }
            Field f = fieldMap.get(index);
            if (Objects.nonNull(f)) {
                throw new IllegalStateException("Index cannot be repeated. Please check it.");
            }
            fieldMap.put(index, field);
        }
        return fieldMap;
    }

    private <T> List<T> getDataFromFile(Sheet sheet, Map<Integer, Field> fieldMap) {
        long startTime = System.currentTimeMillis();
        final int firstRowNum = sheet.getFirstRowNum();
        final int lastRowNum = sheet.getLastRowNum();
        log.info("FirstRowNum:{},LastRowNum:{}", firstRowNum, lastRowNum);
        if (lastRowNum < 0) {
            log.info("Reading excel takes {} milliseconds", System.currentTimeMillis() - startTime);
            return Collections.emptyList();
        }
        DataFormatter formatter = new DataFormatter();
        if (parallelRead) {
            List<ParallelContainer<T>> result = IntStream.rangeClosed(firstRowNum, lastRowNum).parallel().mapToObj(rowNum -> {
                Row row = sheet.getRow(rowNum);
                if (Objects.isNull(row)) {
                    log.info("Row of {} is null,it will be ignored.", rowNum);
                    return null;
                }
                boolean noMatchResult = rowFilter.negate().test(row);
                if (noMatchResult) {
                    log.info("Row of {} does not meet the filtering criteria, it will be ignored.", rowNum);
                    return null;
                }
                int lastColNum = row.getLastCellNum();
                if (lastColNum < 0) {
                    return null;
                }
                T obj = instanceObj(fieldMap, formatter, row);
                return new ParallelContainer<>(rowNum, obj);
            }).filter(Objects::nonNull).collect(Collectors.toList());
            log.info("Reading excel takes {} milliseconds", System.currentTimeMillis() - startTime);
            return result.stream().sorted(Comparator.comparing(ParallelContainer::getIndex)).map(ParallelContainer::getData).collect(Collectors.toList());
        } else {
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
                result.add(obj);
            }
            log.info("Reading excel takes {} milliseconds", System.currentTimeMillis() - startTime);
            return result;
        }
    }

    private <T> void readThenConsume(Sheet sheet, Map<Integer, Field> fieldMap, Consumer<T> consumer) {
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
            consumer.accept(obj);
        }
        log.info("Reading excel takes {} milliseconds", System.currentTimeMillis() - startTime);
    }

    @SuppressWarnings("unchecked")
    private <T> T instanceObj(Map<Integer, Field> fieldMap, DataFormatter formatter, Row row) {
        T obj;
        try {
            obj = (T) dataType.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
        fieldMap.forEach((key, field) -> {
            Cell cell = row.getCell(key, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (Objects.isNull(cell)) {
                return;
            }
            String content = formatter.formatCellValue(cell);
            field.setAccessible(true);
            ReadConverterContext.convert(content, field, obj);
        });
        return obj;
    }
}
