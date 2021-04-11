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
package com.github.liaochong.myexcel.core.parser;

import com.github.liaochong.myexcel.core.strategy.SheetStrategy;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;

/**
 * 解析配置
 *
 * @author liaochong
 * @version 1.0
 */
public class ParseConfig {

    private WidthStrategy widthStrategy;

    private SheetStrategy sheetStrategy;

    private boolean isComputeAutoWidth;

    private final boolean isMultiSheet;

    public ParseConfig(WidthStrategy widthStrategy) {
        this.widthStrategy = widthStrategy;
        this.isComputeAutoWidth = WidthStrategy.isComputeAutoWidth(widthStrategy);
        this.sheetStrategy = SheetStrategy.MULTI_SHEET;
        this.isMultiSheet = true;
    }

    public ParseConfig(WidthStrategy widthStrategy, SheetStrategy sheetStrategy) {
        this.widthStrategy = widthStrategy;
        this.sheetStrategy = sheetStrategy;
        this.isComputeAutoWidth = WidthStrategy.isComputeAutoWidth(widthStrategy);
        this.isMultiSheet = SheetStrategy.isMultiSheet(sheetStrategy);
    }

    public WidthStrategy getWidthStrategy() {
        return this.widthStrategy;
    }

    public boolean isComputeAutoWidth() {
        return this.isComputeAutoWidth;
    }

    public void setWidthStrategy(WidthStrategy widthStrategy) {
        this.widthStrategy = widthStrategy;
    }

    public void setComputeAutoWidth(boolean isComputeAutoWidth) {
        this.isComputeAutoWidth = isComputeAutoWidth;
    }

    public SheetStrategy getSheetStrategy() {
        return this.sheetStrategy;
    }

    public void setSheetStrategy(SheetStrategy sheetStrategy) {
        this.sheetStrategy = sheetStrategy;
    }

    public boolean isMultiSheet() {
        return this.isMultiSheet;
    }
}
