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
package com.github.liaochong.myexcel.utils;

import com.github.liaochong.myexcel.core.Configuration;
import com.github.liaochong.myexcel.core.WorkbookType;
import com.github.liaochong.myexcel.core.annotation.ExcelModel;
import com.github.liaochong.myexcel.core.annotation.ExcelTable;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;

import java.util.Arrays;

/**
 * @author liaochong
 * @version 1.0
 */
public final class ConfigurationUtil {

    public static void parseConfiguration(ClassFieldContainer classFieldContainer, Configuration configuration) {
        ClassFieldContainer parentContainer = classFieldContainer.getParent();
        if (parentContainer != null) {
            parseConfiguration(parentContainer, configuration);
        }
        if (classFieldContainer.getClazz() == Object.class) {
            return;
        }
        ExcelModel excelModel = classFieldContainer.getClazz().getAnnotation(ExcelModel.class);
        if (excelModel == null) {
            ExcelTable excelTable = classFieldContainer.getClazz().getAnnotation(ExcelTable.class);
            if (excelTable == null) {
                return;
            }
            if (!excelTable.sheetName().isEmpty()) {
                configuration.setSheetName(excelTable.sheetName());
            }
            if (!WorkbookType.isNone(excelTable.workbookType())) {
                configuration.setWorkbookType(excelTable.workbookType());
            }
            configuration.setExcludeParent(excelTable.excludeParent());
            configuration.setIncludeAllField(excelTable.includeAllField());
            if (!excelTable.defaultValue().isEmpty()) {
                configuration.setDefaultValue(excelTable.defaultValue());
            }
            configuration.setWrapText(excelTable.wrapText());
            if (!excelTable.titleSeparator().isEmpty()) {
                configuration.setTitleSeparator(excelTable.titleSeparator());
            }
            configuration.setIgnoreStaticFields(excelTable.ignoreStaticFields());
            if (excelTable.titleRowHeight() != -1) {
                configuration.setTitleRowHeight(excelTable.titleRowHeight());
            }
            if (excelTable.rowHeight() != -1) {
                configuration.setRowHeight(excelTable.rowHeight());
            }
            if (excelTable.style().length != 0) {
                configuration.getStyle().addAll(Arrays.asList(excelTable.style()));
            }
            configuration.setUseFieldNameAsTitle(excelTable.useFieldNameAsTitle());
        } else {
            if (!excelModel.sheetName().isEmpty()) {
                configuration.setSheetName(excelModel.sheetName());
            }
            if (!WorkbookType.isNone(excelModel.workbookType())) {
                configuration.setWorkbookType(excelModel.workbookType());
            }
            configuration.setExcludeParent(excelModel.excludeParent());
            configuration.setIncludeAllField(excelModel.includeAllField());
            if (!excelModel.defaultValue().isEmpty()) {
                configuration.setDefaultValue(excelModel.defaultValue());
            }
            configuration.setWrapText(excelModel.wrapText());
            if (!excelModel.titleSeparator().isEmpty()) {
                configuration.setTitleSeparator(excelModel.titleSeparator());
            }
            configuration.setIgnoreStaticFields(excelModel.ignoreStaticFields());
            if (excelModel.titleRowHeight() != -1) {
                configuration.setTitleRowHeight(excelModel.titleRowHeight());
            }
            if (excelModel.rowHeight() != -1) {
                configuration.setRowHeight(excelModel.rowHeight());
            }
            if (excelModel.style().length != 0) {
                configuration.getStyle().addAll(Arrays.asList(excelModel.style()));
            }
            configuration.setUseFieldNameAsTitle(excelModel.useFieldNameAsTitle());
            if (!excelModel.decimalFormat().isEmpty()) {
                configuration.setDecimalFormat(excelModel.decimalFormat());
            }
            if (!excelModel.dateFormat().isEmpty()) {
                configuration.setDateFormat(excelModel.dateFormat());
            }
            if (!excelModel.dateTimeFormat().isEmpty()) {
                configuration.setDateTimeFormat(excelModel.dateTimeFormat());
            }
        }
    }
}
