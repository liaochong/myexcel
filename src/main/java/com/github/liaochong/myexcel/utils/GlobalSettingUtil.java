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

import com.github.liaochong.myexcel.core.GlobalSetting;
import com.github.liaochong.myexcel.core.annotation.ExcelModel;
import com.github.liaochong.myexcel.core.annotation.ExcelTable;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;

import java.util.Arrays;

/**
 * @author liaochong
 * @version 1.0
 */
public final class GlobalSettingUtil {

    public static void setGlobalSetting(ClassFieldContainer classFieldContainer, GlobalSetting globalSetting) {
        ClassFieldContainer parentContainer = classFieldContainer.getParent();
        if (parentContainer != null) {
            setGlobalSetting(parentContainer, globalSetting);
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
                globalSetting.setSheetName(excelTable.sheetName());
            }
            if (!globalSetting.isFixedWorkbookType()) {
                globalSetting.setWorkbookType(excelTable.workbookType());
            }
            globalSetting.setExcludeParent(excelTable.excludeParent());
            globalSetting.setIncludeAllField(excelTable.includeAllField());
            if (!excelTable.defaultValue().isEmpty()) {
                globalSetting.setDefaultValue(excelTable.defaultValue());
            }
            globalSetting.setWrapText(excelTable.wrapText());
            if (!excelTable.titleSeparator().isEmpty()) {
                globalSetting.setTitleSeparator(excelTable.titleSeparator());
            }
            globalSetting.setIgnoreStaticFields(excelTable.ignoreStaticFields());
            if (excelTable.titleRowHeight() != -1) {
                globalSetting.setTitleRowHeight(excelTable.titleRowHeight());
            }
            if (excelTable.rowHeight() != -1) {
                globalSetting.setRowHeight(excelTable.rowHeight());
            }
            if (excelTable.style().length != 0) {
                globalSetting.getGlobalStyle().addAll(Arrays.asList(excelTable.style()));
            }
            globalSetting.setUseFieldNameAsTitle(excelTable.useFieldNameAsTitle());
        } else {
            if (!excelModel.sheetName().isEmpty()) {
                globalSetting.setSheetName(excelModel.sheetName());
            }
            if (!globalSetting.isFixedWorkbookType()) {
                globalSetting.setWorkbookType(excelModel.workbookType());
            }
            globalSetting.setExcludeParent(excelModel.excludeParent());
            globalSetting.setIncludeAllField(excelModel.includeAllField());
            if (!excelModel.defaultValue().isEmpty()) {
                globalSetting.setDefaultValue(excelModel.defaultValue());
            }
            globalSetting.setWrapText(excelModel.wrapText());
            if (!excelModel.titleSeparator().isEmpty()) {
                globalSetting.setTitleSeparator(excelModel.titleSeparator());
            }
            globalSetting.setIgnoreStaticFields(excelModel.ignoreStaticFields());
            if (excelModel.titleRowHeight() != -1) {
                globalSetting.setTitleRowHeight(excelModel.titleRowHeight());
            }
            if (excelModel.rowHeight() != -1) {
                globalSetting.setRowHeight(excelModel.rowHeight());
            }
            if (excelModel.style().length != 0) {
                globalSetting.getGlobalStyle().addAll(Arrays.asList(excelModel.style()));
            }
            globalSetting.setUseFieldNameAsTitle(excelModel.useFieldNameAsTitle());
            if (!excelModel.decimalFormat().isEmpty()) {
                globalSetting.setDecimalFormat(excelModel.decimalFormat());
            }
            if (!excelModel.dateFormat().isEmpty()) {
                globalSetting.setDateFormat(excelModel.dateFormat());
            }
            if (!excelModel.dateTimeFormat().isEmpty()) {
                globalSetting.setDateTimeFormat(excelModel.dateTimeFormat());
            }
        }
    }
}
