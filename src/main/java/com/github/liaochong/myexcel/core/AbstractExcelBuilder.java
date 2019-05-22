/*
 * Copyright 2017 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.io.TempFileOperator;
import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;

import java.util.Objects;

/**
 * excel创建者接口
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public abstract class AbstractExcelBuilder implements ExcelBuilder {

    protected HtmlToExcelFactory htmlToExcelFactory = new HtmlToExcelFactory();

    protected TempFileOperator tempFileOperator = new TempFileOperator();

    @Override
    public AbstractExcelBuilder workbookType(@NonNull WorkbookType workbookType) {
        htmlToExcelFactory.workbookType(workbookType);
        if (WorkbookType.isSxlsx(workbookType)) {
            autoWidthStrategy(AutoWidthStrategy.NO_AUTO);
        }
        return this;
    }

    @Override
    public AbstractExcelBuilder rowAccessWindowSize(int rowAccessWindowSize) {
        htmlToExcelFactory.rowAccessWindowSize(rowAccessWindowSize);
        return this;
    }

    @Override
    public AbstractExcelBuilder useDefaultStyle() {
        htmlToExcelFactory.useDefaultStyle();
        return this;
    }

    @Override
    public AbstractExcelBuilder autoWidthStrategy(@NonNull AutoWidthStrategy autoWidthStrategy) {
        htmlToExcelFactory.autoWidthStrategy(autoWidthStrategy);
        return this;
    }

    @Override
    public AbstractExcelBuilder freezePanes(FreezePane... freezePanes) {
        if (Objects.isNull(freezePanes) || freezePanes.length == 0) {
            return this;
        }
        htmlToExcelFactory.freezePanes(freezePanes);
        return this;
    }
}
