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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author liaochong
 * @version 1.0
 */
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
        Workbook wb;
        if (StringUtil.isBlank(password)) {
            wb = WorkbookFactory.create(fileInputStream);
        } else {
            wb = WorkbookFactory.create(fileInputStream, password);
        }
        Sheet sheet = wb.getSheetAt(sheetIndex);
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
        Workbook wb;
        if (StringUtil.isBlank(password)) {
            wb = WorkbookFactory.create(file);
        } else {
            wb = WorkbookFactory.create(file, password);
        }
        Sheet sheet = wb.getSheetAt(sheetIndex);
        return getDataFromFile(sheet, fieldMap);
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

    @SuppressWarnings("unchecked")
    private <T> List<T> getDataFromFile(Sheet sheet, Map<Integer, Field> fieldMap) {
        final int firstRowNum = sheet.getFirstRowNum();
        final int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return Collections.emptyList();
        }
        DataFormatter formatter = new DataFormatter();
        if (parallelRead) {
            List<ParallelContainer<T>> result = IntStream.rangeClosed(firstRowNum, lastRowNum).parallel().mapToObj(rowNum -> {
                Row row = sheet.getRow(rowNum);
                if (Objects.isNull(row)) {
                    return null;
                }
                boolean noMatchResult = rowFilter.negate().test(row);
                if (noMatchResult) {
                    return null;
                }
                int lastColNum = row.getLastCellNum();
                if (lastColNum < 0) {
                    return null;
                }
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

                return new ParallelContainer<>(rowNum, obj);
            }).filter(Objects::nonNull).collect(Collectors.toList());

            return result.stream().sorted(Comparator.comparing(ParallelContainer::getIndex)).map(ParallelContainer::getData).collect(Collectors.toList());
        } else {
            List<T> result = new ArrayList<>(fieldMap.size());
            for (int i = firstRowNum; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if (Objects.isNull(row)) {
                    continue;
                }
                boolean noMatchResult = rowFilter.negate().test(row);
                if (noMatchResult) {
                    continue;
                }
                int lastColNum = row.getLastCellNum();
                if (lastColNum < 0) {
                    continue;
                }
                T obj;
                try {
                    obj = (T) dataType.newInstance();
                } catch (InstantiationException | IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
                result.add(obj);
                fieldMap.forEach((key, field) -> {
                    Cell cell = row.getCell(key, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (Objects.isNull(cell)) {
                        return;
                    }
                    String content = formatter.formatCellValue(cell);
                    field.setAccessible(true);
                    ReadConverterContext.convert(content, field, obj);
                });
            }
            return result;
        }
    }
}
