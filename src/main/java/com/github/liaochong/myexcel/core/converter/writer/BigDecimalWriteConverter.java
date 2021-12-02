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
package com.github.liaochong.myexcel.core.converter.writer;

import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.ConvertContext;
import com.github.liaochong.myexcel.core.converter.WriteConverter;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;

/**
 * bigdecimal格式化
 *
 * @author liaochong
 * @version 1.0
 */
public class BigDecimalWriteConverter implements WriteConverter {

    @Override
    public boolean support(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        return fieldType == BigDecimal.class;
    }

    @Override
    public Pair<Class, Object> convert(Field field, Class<?> fieldType, Object fieldVal, ConvertContext convertContext) {
        if (!convertContext.isConvertCsv()) {
            return Pair.of(Double.class, ((BigDecimal) fieldVal).toPlainString());
        }
        ExcelColumnMapping excelColumnMapping = convertContext.getExcelColumnMappingMap().get(field);
        String format = convertContext.getConfiguration().decimalFormat;
        if (excelColumnMapping != null && !excelColumnMapping.getFormat().isEmpty()) {
            format = excelColumnMapping.getFormat();
        }
        if (format.isEmpty()) {
            return Pair.of(Double.class, ((BigDecimal) fieldVal).toPlainString());
        }
        String[] formatSplits = format.split("\\.");
        BigDecimal value = (BigDecimal) fieldVal;
        if (formatSplits.length == 2) {
            value = value.setScale(formatSplits[1].length(), RoundingMode.HALF_UP);
        }
        DecimalFormat decimalFormat = new DecimalFormat(format);
        return Pair.of(String.class, decimalFormat.format(value));
    }
}
