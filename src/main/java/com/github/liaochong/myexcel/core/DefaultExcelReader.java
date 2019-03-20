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
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public class DefaultExcelReader {

    private static final ReadConverterContext READ_CONVERTER_CONTEXT = ReadConverterContext.getInstance();

    private Class<?> dataType;

    private int sheetIndex = 0;

    private DefaultExcelReader(Class<?> dataType) {
        this.dataType = dataType;
    }

    public static <T> DefaultExcelReader of(@NonNull Class<T> clazz) {
        return new DefaultExcelReader(clazz);
    }

    public DefaultExcelReader sheet(int index) {
        this.sheetIndex = index;
        return this;
    }

    public <T> List<T> read(@NonNull File file) {
        if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls")) {
            throw new IllegalArgumentException();
        }
        List<Field> sortedFields = getSortedField();
        if (file.getName().endsWith(".xlsx")) {
            try (OPCPackage pkg = OPCPackage.open(file)) {
                Workbook wb = new XSSFWorkbook(pkg);
                Sheet sheet = wb.getSheetAt(sheetIndex);
                return getDataFromFile(sheet, sortedFields);
            } catch (IOException | InvalidFormatException | IllegalAccessException | InstantiationException e) {
                throw new RuntimeException(e);
            }
        }
        try (POIFSFileSystem fs = new POIFSFileSystem(file)) {
            Workbook wb = new HSSFWorkbook(fs.getRoot(), true);
            Sheet sheet = wb.getSheetAt(sheetIndex);
            return getDataFromFile(sheet, sortedFields);
        } catch (IOException | IllegalAccessException | InstantiationException e) {
            throw new RuntimeException(e);
        }
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
                READ_CONVERTER_CONTEXT.convert(content, field, obj);
            }
        }
        return result;
    }
}
