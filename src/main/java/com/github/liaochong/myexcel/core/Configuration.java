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
    public String sheetName;
    /**
     * The type of workbook to be built
     */
    public WorkbookType workbookType = WorkbookType.SXLSX;
    /**
     * 宽度策略
     */
    public WidthStrategy widthStrategy;
    /**
     * 是否排除父类字段
     */
    public boolean excludeParent = false;
    /**
     * 是否导出所有字段，否，则只导出含{@link com.github.liaochong.myexcel.core.annotation.ExcelColumn}注解字段
     */
    public boolean includeAllField = true;
    /**
     * 当对应字段的值为null时所需要替换的默认值
     */
    public String defaultValue;
    /**
     * 是否自动换行
     */
    public boolean wrapText = true;
    /**
     * 多级标题所需的分离标志
     */
    public String titleSeparator = Constants.ARROW;
    /**
     * 是否忽略静态字段
     */
    public boolean ignoreStaticFields = true;
    /**
     * 标题行高度
     */
    public int titleRowHeight;
    /**
     * 内容行高度
     */
    public int rowHeight;
    /**
     * 全局样式
     */
    public Set<String> style = new HashSet<>();
    /**
     * 是否使用字段名称作为标题，当{@link com.github.liaochong.myexcel.core.annotation.ExcelColumn}设定了title，则覆盖
     */
    public boolean useFieldNameAsTitle = false;
    /**
     * LocalDate类型数据全局格式化
     */
    public String dateFormat = "yyyy-MM-dd";
    /**
     * Date、LocalDateTime类型数据全局格式化
     */
    public String dateTimeFormat = "yyyy-MM-dd HH:mm:ss";
    /**
     * LocalTime格式化
     */
    public String localTimeFormat = "HH:mm:ss";
    /**
     * 数值类全局格式化
     */
    public String decimalFormat = "";

    public boolean computeAutoWidth;

    /**
     * 绑定的上下文，适用spring等容器环境
     */
    public Map<Class<?>, Object> applicationBeans = Collections.emptyMap();

    public void setWidthStrategy(WidthStrategy widthStrategy) {
        this.widthStrategy = widthStrategy;
        this.computeAutoWidth = WidthStrategy.isComputeAutoWidth(widthStrategy);
    }
}
