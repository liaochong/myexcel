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
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Predicate;

/**
 * excel读取抽象
 *
 * @author liaochong
 * @version 1.0
 */
public abstract class AbstractExcelReader<T> implements ExcelReader<T> {

    protected static final int DEFAULT_SHEET_INDEX = 0;

    protected Class<T> dataType;

    protected int sheetIndex = DEFAULT_SHEET_INDEX;

    protected Predicate<Row> rowFilter = row -> true;

    protected Map<Integer, Field> getFieldMap() {
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
            field.setAccessible(true);
            fieldMap.put(index, field);
        }
        return fieldMap;
    }


}
