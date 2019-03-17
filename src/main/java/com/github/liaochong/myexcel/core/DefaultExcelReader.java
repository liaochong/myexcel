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
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import lombok.NonNull;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
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

    private Class<?> dataType;

    private DefaultExcelReader(Class<?> dataType) {
        this.dataType = dataType;
    }

    public static <T> DefaultExcelReader of(@NonNull Class<T> clazz) {
        return new DefaultExcelReader(clazz);
    }

    public <T> List<T> read(@NonNull File file) throws Exception {
        if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls")) {
            throw new IllegalArgumentException();
        }
        List<Field> sortedFields = getSortedField();
        Workbook wb = null;
        Sheet sheet;
        List<T> result = null;
        if (file.getName().endsWith(".xlsx")) {
            OPCPackage pkg = null;
            try {
                pkg = OPCPackage.open(file);
                wb = new XSSFWorkbook(pkg);
                sheet = wb.getSheetAt(0);
                result = getDataFromFile(sheet, sortedFields);
            } catch (IOException | InvalidFormatException e) {
                e.printStackTrace();
            } finally {
                if (Objects.nonNull(pkg)) {
                    pkg.close();
                }
            }
        } else if (file.getName().endsWith(".xls")) {
            POIFSFileSystem fs = null;
            try {
                fs = new POIFSFileSystem(file);
                wb = new HSSFWorkbook(fs.getRoot(), true);
                sheet = wb.getSheetAt(0);
                result = getDataFromFile(sheet, sortedFields);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                if (Objects.nonNull(fs)) {
                    fs.close();
                }
            }
        }
        return result;
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
                Object content = null;
                switch (cell.getCellType()) {
                    case STRING:
                        content = cell.getRichStringCellValue().getString();
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            content = cell.getDateCellValue();
                        } else {
                            content = cell.getNumericCellValue();
                        }
                        break;
                    case BOOLEAN:
                        content = cell.getBooleanCellValue();
                        break;
                    case FORMULA:
                        content = cell.getCellFormula();
                        break;
                    case BLANK:
                        break;
                    default:
                }
                sortedFields.get(j).set(obj, content);
            }
        }
        return result;
    }
}
