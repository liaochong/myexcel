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
package com.github.liaochong.myexcel.core.converter;

import com.github.liaochong.myexcel.core.Configuration;
import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.constant.AllConverter;
import com.github.liaochong.myexcel.core.constant.CsvConverter;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
public class ConvertContext {
    /**
     * {@link com.github.liaochong.myexcel.core.annotation.ExcelModel} setting
     */
    public Configuration configuration = new Configuration();

    /**
     * {@link com.github.liaochong.myexcel.core.annotation.ExcelColumn} mapping
     */
    public Map<Field, ExcelColumnMapping> excelColumnMappingMap = new HashMap<>();
    /**
     * csv or excel
     */
    public Class converterType;

    public boolean isConvertCsv;

    public ConvertContext(boolean isConvertCsv) {
        this.isConvertCsv = isConvertCsv;
        this.converterType = isConvertCsv ? CsvConverter.class : AllConverter.class;
    }
}
