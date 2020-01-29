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
class GlobalSetting {
    /**
     * The name of the sheet to be built
     */
    String sheetName;
    /**
     * The type of workbook to be built
     */
    WorkbookType workbookType;

    WidthStrategy widthStrategy;

    boolean excludeParent = false;

    boolean includeAllField = true;

    String defaultValue;

    boolean wrapText = true;

    String titleSeparator = Constants.ARROW;

    boolean ignoreStaticFields = true;

    int titleRowHeight;

    int rowHeight;

    Set<String> globalStyle = new HashSet<>();

    boolean useFieldNameAsTitle = false;

    String dateFormat;

    String decimalFormat;
}
