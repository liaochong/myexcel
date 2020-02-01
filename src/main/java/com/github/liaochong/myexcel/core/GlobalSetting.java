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
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

import java.util.HashSet;
import java.util.Set;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class GlobalSetting {
    /**
     * 方法设定时固定-sheet名称，注解对应属性将无法变更
     */
    boolean fixedSheetName;
    /**
     * 方法设定时固定-工作簿类型，注解对应属性将无法变更
     */
    boolean fixedWorkbookType;
    /**
     * 方法设定时固定-宽度策略，注解对应属性将无法变更
     */
    boolean fixedWidthStrategy;
    /**
     * 方法设定时固定-全局样式，注解对应属性将无法变更
     */
    boolean fixedGlobalStyle;
    /**
     * The name of the sheet to be built
     */
    String sheetName;
    /**
     * The type of workbook to be built
     */
    WorkbookType workbookType = WorkbookType.SXLSX;
    /**
     * 宽度策略
     */
    WidthStrategy widthStrategy;
    /**
     * 是否排除父类字段
     */
    boolean excludeParent = false;
    /**
     * 是否导出所有字段，否，则只导出含{@link com.github.liaochong.myexcel.core.annotation.ExcelColumn}注解字段
     */
    boolean includeAllField = true;
    /**
     * 当对应字段的值为null时所需要替换的默认值
     */
    String defaultValue;
    /**
     * 是否自动换行
     */
    boolean wrapText = true;
    /**
     * 多级标题所需的分离标志
     */
    String titleSeparator = Constants.ARROW;
    /**
     * 是否忽略静态字段
     */
    boolean ignoreStaticFields = true;
    /**
     * 标题行高度
     */
    int titleRowHeight;
    /**
     * 内容行高度
     */
    int rowHeight;
    /**
     * 全局样式
     */
    Set<String> globalStyle = new HashSet<>();
    /**
     * 是否使用字段名称作为标题，当{@link com.github.liaochong.myexcel.core.annotation.ExcelColumn}设定了title，则覆盖
     */
    boolean useFieldNameAsTitle = false;
    /**
     * LocalDate类型数据全局格式化
     */
    String dateFormat = "yyyy-MM-dd";
    /**
     * Date、LocalDateTime类型数据全局格式化
     */
    String dateTimeFormat = "yyyy-MM-dd HH:mm:ss";
    /**
     * 数值类全局格式化
     */
    String decimalFormat;
}
