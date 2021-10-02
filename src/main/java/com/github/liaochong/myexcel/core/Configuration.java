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

import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;

import java.util.Collections;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * @author liaochong
 * @version 1.0
 */
public class Configuration {
    /**
     * The name of the sheet to be built
     */
    private String sheetName;
    /**
     * The type of workbook to be built
     */
    private WorkbookType workbookType = WorkbookType.SXLSX;
    /**
     * 宽度策略
     */
    private WidthStrategy widthStrategy;
    /**
     * 是否排除父类字段
     */
    private boolean excludeParent = false;
    /**
     * 是否导出所有字段，否，则只导出含{@link com.github.liaochong.myexcel.core.annotation.ExcelColumn}注解字段
     */
    private boolean includeAllField = true;
    /**
     * 当对应字段的值为null时所需要替换的默认值
     */
    private String defaultValue;
    /**
     * 是否自动换行
     */
    private boolean wrapText = true;
    /**
     * 多级标题所需的分离标志
     */
    private String titleSeparator = Constants.ARROW;
    /**
     * 是否忽略静态字段
     */
    private boolean ignoreStaticFields = true;
    /**
     * 标题行高度
     */
    private int titleRowHeight;
    /**
     * 内容行高度
     */
    private int rowHeight;
    /**
     * 全局样式
     */
    private Set<String> style = new HashSet<>();
    /**
     * 是否使用字段名称作为标题，当{@link com.github.liaochong.myexcel.core.annotation.ExcelColumn}设定了title，则覆盖
     */
    private boolean useFieldNameAsTitle = false;
    /**
     * LocalDate类型数据全局格式化
     */
    private String dateFormat = "yyyy-MM-dd";
    /**
     * Date、LocalDateTime类型数据全局格式化
     */
    private String dateTimeFormat = "yyyy-MM-dd HH:mm:ss";
    /**
     * LocalTime格式化
     */
    private String localTimeFormat = "HH:mm:ss";
    /**
     * 数值类全局格式化
     */
    private String decimalFormat = "";

    private boolean computeAutoWidth;

    /**
     * 绑定的上下文，适用spring等容器环境
     */
    private Map<Class<?>, Object> applicationBeans = Collections.emptyMap();

    public void setWidthStrategy(WidthStrategy widthStrategy) {
        this.widthStrategy = widthStrategy;
        this.computeAutoWidth = WidthStrategy.isComputeAutoWidth(widthStrategy);
    }

    public String getSheetName() {
        return this.sheetName;
    }

    public WorkbookType getWorkbookType() {
        return this.workbookType;
    }

    public WidthStrategy getWidthStrategy() {
        return this.widthStrategy;
    }

    public boolean isExcludeParent() {
        return this.excludeParent;
    }

    public boolean isIncludeAllField() {
        return this.includeAllField;
    }

    public String getDefaultValue() {
        return this.defaultValue;
    }

    public boolean isWrapText() {
        return this.wrapText;
    }

    public String getTitleSeparator() {
        return this.titleSeparator;
    }

    public boolean isIgnoreStaticFields() {
        return this.ignoreStaticFields;
    }

    public int getTitleRowHeight() {
        return this.titleRowHeight;
    }

    public int getRowHeight() {
        return this.rowHeight;
    }

    public Set<String> getStyle() {
        return this.style;
    }

    public boolean isUseFieldNameAsTitle() {
        return this.useFieldNameAsTitle;
    }

    public String getDateFormat() {
        return this.dateFormat;
    }

    public String getDateTimeFormat() {
        return this.dateTimeFormat;
    }

    public String getDecimalFormat() {
        return this.decimalFormat;
    }

    public boolean isComputeAutoWidth() {
        return this.computeAutoWidth;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setWorkbookType(WorkbookType workbookType) {
        this.workbookType = workbookType;
    }

    public void setExcludeParent(boolean excludeParent) {
        this.excludeParent = excludeParent;
    }

    public void setIncludeAllField(boolean includeAllField) {
        this.includeAllField = includeAllField;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public void setWrapText(boolean wrapText) {
        this.wrapText = wrapText;
    }

    public void setTitleSeparator(String titleSeparator) {
        this.titleSeparator = titleSeparator;
    }

    public void setIgnoreStaticFields(boolean ignoreStaticFields) {
        this.ignoreStaticFields = ignoreStaticFields;
    }

    public void setTitleRowHeight(int titleRowHeight) {
        this.titleRowHeight = titleRowHeight;
    }

    public void setRowHeight(int rowHeight) {
        this.rowHeight = rowHeight;
    }

    public void setStyle(Set<String> style) {
        this.style = style;
    }

    public void setUseFieldNameAsTitle(boolean useFieldNameAsTitle) {
        this.useFieldNameAsTitle = useFieldNameAsTitle;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public void setDateTimeFormat(String dateTimeFormat) {
        this.dateTimeFormat = dateTimeFormat;
    }

    public void setDecimalFormat(String decimalFormat) {
        this.decimalFormat = decimalFormat;
    }

    public void setComputeAutoWidth(boolean computeAutoWidth) {
        this.computeAutoWidth = computeAutoWidth;
    }

    public Map<Class<?>, Object> getApplicationBeans() {
        return applicationBeans;
    }

    public void setApplicationBeans(Map<Class<?>, Object> applicationBeans) {
        this.applicationBeans = applicationBeans;
    }

    public String getLocalTimeFormat() {
        return localTimeFormat;
    }

    public void setLocalTimeFormat(String localTimeFormat) {
        this.localTimeFormat = localTimeFormat;
    }
}
