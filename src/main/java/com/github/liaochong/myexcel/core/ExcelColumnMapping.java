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
import com.github.liaochong.myexcel.core.constant.FileType;
import com.github.liaochong.myexcel.core.constant.LinkType;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.experimental.FieldDefaults;

/**
 * @author liaochong
 * @version 1.0
 */
@Getter
@FieldDefaults(level = AccessLevel.PRIVATE)
public final class ExcelColumnMapping {

    /**
     * 列标题
     */
    String title;

    /**
     * 顺序，数值越大越靠后
     */
    int order;

    /**
     * 列索引，从零开始，不允许重复
     */
    int index;

    /**
     * 分组
     */
    Class<?>[] groups;

    /**
     * 为null时默认值
     */
    String defaultValue;

    /**
     * 宽度
     */
    int width;

    /**
     * 是否强制转换成字符串
     */
    boolean convertToString;

    /**
     * 格式化，时间、金额等
     */
    String format;

    /**
     * 样式
     */
    String[] style;

    /**
     * 链接
     */
    LinkType linkType;

    /**
     * 简单映射，如"1:男,2:女"
     */
    String mapping;

    /**
     * 文件类型
     */
    FileType fileType;

    public static ExcelColumnMapping mapping(ExcelColumn excelColumn) {
        ExcelColumnMapping result = new ExcelColumnMapping();
        result.title = excelColumn.title();
        result.order = excelColumn.order();
        result.index = excelColumn.index();
        result.groups = excelColumn.groups();
        result.defaultValue = excelColumn.defaultValue();
        result.width = excelColumn.width();
        result.convertToString = excelColumn.convertToString();
        if (!excelColumn.format().isEmpty()) {
            result.format = excelColumn.format();
        } else if (!excelColumn.dateFormatPattern().isEmpty()) {
            result.format = excelColumn.dateFormatPattern();
        } else if (!excelColumn.decimalFormat().isEmpty()) {
            result.format = excelColumn.decimalFormat();
        } else {
            result.format = "";
        }
        result.style = excelColumn.style();
        result.linkType = excelColumn.linkType();
        result.mapping = excelColumn.mapping();
        result.fileType = excelColumn.fileType();
        return result;
    }
}
