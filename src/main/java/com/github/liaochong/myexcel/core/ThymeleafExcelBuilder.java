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

import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.core.strategy.SheetStrategy;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.core.templatehandler.ThymeleafTemplateHandler;

/**
 * ThymeleafExcelBuilder
 *
 * @author liaochong
 * @version 1.0
 */
public class ThymeleafExcelBuilder extends AbstractExcelBuilder {

    public ThymeleafExcelBuilder() {
        super(ThymeleafTemplateHandler.class);
    }

    @Override
    public AbstractExcelBuilder useDefaultStyle() {
        htmlToExcelFactory.useDefaultStyle();
        return this;
    }

    @Override
    public ExcelBuilder applyDefaultStyle() {
        htmlToExcelFactory.applyDefaultStyle();
        return this;
    }

    @Override
    public AbstractExcelBuilder widthStrategy(WidthStrategy widthStrategy) {
        htmlToExcelFactory.widthStrategy(widthStrategy);
        return this;
    }

    @Deprecated
    @Override
    public AbstractExcelBuilder autoWidthStrategy(AutoWidthStrategy autoWidthStrategy) {
        htmlToExcelFactory.widthStrategy(AutoWidthStrategy.map(autoWidthStrategy));
        return this;
    }

    @Override
    public AbstractExcelBuilder sheetStrategy(SheetStrategy sheetStrategy) {
        htmlToExcelFactory.sheetStrategy(sheetStrategy);
        return this;
    }
}
