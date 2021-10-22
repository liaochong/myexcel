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
import com.github.liaochong.myexcel.core.annotation.Prompt;
import com.github.liaochong.myexcel.core.constant.FileType;
import com.github.liaochong.myexcel.core.constant.LinkType;
import com.github.liaochong.myexcel.core.converter.CustomWriteConverter;
import com.github.liaochong.myexcel.utils.StringUtil;

/**
 * @author liaochong
 * @version 1.0
 */
public final class ExcelColumnMapping {

    /**
     * 列标题
     */
    private String title;

    /**
     * 顺序，数值越大越靠后
     */
    private int order;

    /**
     * 列索引，从零开始，不允许重复
     */
    private int index;

    /**
     * 分组
     */
    private Class<?>[] groups;

    /**
     * 为null时默认值
     */
    private String defaultValue;

    /**
     * 宽度
     */
    private int width;

    /**
     * 是否强制转换成字符串
     */
    private boolean convertToString;

    /**
     * 格式化，时间、金额等
     */
    private String format;

    /**
     * 样式
     */
    private String[] style;

    /**
     * 链接
     */
    private LinkType linkType;

    /**
     * 简单映射，如"1:男,2:女"
     */
    private String mapping;
    /**
     * 自定义写转换器
     */
    private Class<? extends CustomWriteConverter> customWriteConverter;

    /**
     * 文件类型
     */
    private FileType fileType;

    /**
     * 是否为公式
     */
    private boolean formula;

    /**
     * 提示语
     */
    private PromptContainer promptContainer;

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
        result.formula = excelColumn.formula();
        result.customWriteConverter = excelColumn.writeConverter();
        // 提示
        Prompt prompt = excelColumn.prompt();
        if (StringUtil.isNotBlank(prompt.text())) {
            PromptContainer promptContainer = new PromptContainer();
            promptContainer.setTitle(prompt.title());
            promptContainer.setText(prompt.text());
            result.promptContainer = promptContainer;
        }
        return result;
    }

    public String getTitle() {
        return this.title;
    }

    public int getOrder() {
        return this.order;
    }

    public int getIndex() {
        return this.index;
    }

    public Class<?>[] getGroups() {
        return this.groups;
    }

    public String getDefaultValue() {
        return this.defaultValue;
    }

    public int getWidth() {
        return this.width;
    }

    public boolean isConvertToString() {
        return this.convertToString;
    }

    public String getFormat() {
        return this.format;
    }

    public String[] getStyle() {
        return this.style;
    }

    public LinkType getLinkType() {
        return this.linkType;
    }

    public String getMapping() {
        return this.mapping;
    }

    public FileType getFileType() {
        return this.fileType;
    }

    public boolean isFormula() {
        return this.formula;
    }

    public PromptContainer getPromptContainer() {
        return promptContainer;
    }

    public void setPromptContainer(PromptContainer promptContainer) {
        this.promptContainer = promptContainer;
    }

    public Class<? extends CustomWriteConverter> getCustomWriteConverter() {
        return customWriteConverter;
    }

    public void setCustomWriteConverter(Class<? extends CustomWriteConverter> customWriteConverter) {
        this.customWriteConverter = customWriteConverter;
    }
}
