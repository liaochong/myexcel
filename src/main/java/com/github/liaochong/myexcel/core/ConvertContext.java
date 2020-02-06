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

import com.github.liaochong.myexcel.core.constant.CsvConverter;
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

import java.lang.reflect.Field;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class ConvertContext {

    GlobalSetting globalSetting;

    Map<Field, ExcelColumnMapping> excelColumnMappingMap;
    /**
     * csv or excel
     */
    Class converterType;

    boolean isConvertCsv;

    public ConvertContext(GlobalSetting globalSetting, Map<Field, ExcelColumnMapping> excelColumnMappingMap) {
        this.globalSetting = globalSetting;
        this.excelColumnMappingMap = excelColumnMappingMap;
    }

    public void setConverterType(Class converterType) {
        this.converterType = converterType;
        this.isConvertCsv = CsvConverter.class == converterType;
    }
}
