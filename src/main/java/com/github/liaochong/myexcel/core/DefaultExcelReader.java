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
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import lombok.NonNull;
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
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public class DefaultExcelReader {

    private static final int DEFAULT_SHEET_INDEX = 0;

    private Class<?> dataType;

    private int sheetIndex = DEFAULT_SHEET_INDEX;

    private Predicate<Row> rowFilter = row -> true;

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

    public <T> List<T> read(@NonNull InputStream inputStream) throws IOException, IllegalAccessException, InstantiationException {
        List<Field> sortedFields = getSortedField();
        Workbook wb = WorkbookFactory.create(inputStream);
        Sheet sheet = wb.getSheetAt(sheetIndex);
        return getDataFromFile(sheet, sortedFields);
    }

    public <T> List<T> read(@NonNull File file) throws IOException, IllegalAccessException, InstantiationException {
        if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls")) {
            throw new IllegalArgumentException();
        }
        List<Field> sortedFields = getSortedField();
        Workbook wb = WorkbookFactory.create(file);
        Sheet sheet = wb.getSheetAt(sheetIndex);
        return getDataFromFile(sheet, sortedFields);
    }

    private List<Field> getSortedField() {
        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
        return classFieldContainer.getFieldsByAnnotation(ExcelColumn.class).stream().sorted((field1, field2) -> {
            ExcelColumn excelColumn1 = field1.getAnnotation(ExcelColumn.class);
            ExcelColumn excelColumn2 = field2.getAnnotation(ExcelColumn.class);
            int order1 = excelColumn1.order();
            int order2 = excelColumn2.order();
            if (Objects.equals(order1, order2)) {
                return 0;
            }
            return order1 > order2 ? 1 : -1;
        }).collect(Collectors.toList());
    }

    private <T> List<T> getDataFromFile(Sheet sheet, List<Field> sortedFields) throws InstantiationException, IllegalAccessException {
        final int firstRowNum = sheet.getFirstRowNum();
        final int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return Collections.emptyList();
        }
        DataFormatter formatter = new DataFormatter();
        List<T> result = new ArrayList<>(lastRowNum);
        for (int i = firstRowNum; i < lastRowNum; i++) {
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
            T obj = (T) dataType.newInstance();
            result.add(obj);

            for (int j = 0; j < lastColNum; j++) {
                Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (Objects.isNull(cell)) {
                    continue;
                }
                String content = formatter.formatCellValue(cell);
                Field field = sortedFields.get(j);
                field.setAccessible(true);
                ReadConverterContext.convert(content, field, obj);
            }
        }
        return result;
    }
}
