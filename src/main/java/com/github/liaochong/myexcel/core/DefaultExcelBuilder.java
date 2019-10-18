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

import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
import java.util.List;

/**
 * 默认excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelBuilder implements Closeable {

    private DefaultStreamExcelBuilder streamExcelBuilder;

    private DefaultExcelBuilder(DefaultStreamExcelBuilder streamExcelBuilder) {
        streamExcelBuilder.hasStyle();
        streamExcelBuilder.widthStrategy(WidthStrategy.COMPUTE_AUTO_WIDTH);
        this.streamExcelBuilder = streamExcelBuilder;
    }

    /**
     * 获取实例
     *
     * @return DefaultExcelBuilder
     */
    public static DefaultExcelBuilder getInstance() {
        DefaultExcelBuilder defaultExcelBuilder = new DefaultExcelBuilder(DefaultStreamExcelBuilder.getInstance());
        defaultExcelBuilder.streamExcelBuilder.workbookType(WorkbookType.XLSX);
        return defaultExcelBuilder;
    }

    /**
     * 获取实例
     *
     * @return DefaultExcelBuilder
     */
    public static DefaultExcelBuilder getInstance(Workbook workbook) {
        return new DefaultExcelBuilder(DefaultStreamExcelBuilder.getInstance(workbook));
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @return DefaultExcelBuilder
     */
    public static DefaultExcelBuilder of(@NonNull Class<?> dataType) {
        DefaultExcelBuilder defaultExcelBuilder = new DefaultExcelBuilder(DefaultStreamExcelBuilder.of(dataType));
        defaultExcelBuilder.streamExcelBuilder.workbookType(WorkbookType.XLSX);
        return defaultExcelBuilder;
    }

    public static DefaultExcelBuilder of(@NonNull Class<?> dataType, @NonNull Workbook workbook) {
        return new DefaultExcelBuilder(DefaultStreamExcelBuilder.of(dataType, workbook));
    }

    public DefaultExcelBuilder titles(@NonNull List<String> titles) {
        streamExcelBuilder.titles(titles);
        return this;
    }

    public DefaultExcelBuilder sheetName(@NonNull String sheetName) {
        streamExcelBuilder.sheetName(sheetName);
        return this;
    }

    public DefaultExcelBuilder fieldDisplayOrder(@NonNull List<String> fieldDisplayOrder) {
        streamExcelBuilder.fieldDisplayOrder(fieldDisplayOrder);
        return this;
    }

    public DefaultExcelBuilder workbookType(@NonNull WorkbookType workbookType) {
        streamExcelBuilder.workbookType(workbookType);
        return this;
    }

    public DefaultExcelBuilder noStyle() {
        streamExcelBuilder.noStyle();
        return this;
    }

    public DefaultExcelBuilder widthStrategy(WidthStrategy widthStrategy) {
        streamExcelBuilder.widthStrategy(widthStrategy);
        return this;
    }

    @Deprecated
    public DefaultExcelBuilder autoWidthStrategy(@NonNull AutoWidthStrategy autoWidthStrategy) {
        streamExcelBuilder.autoWidthStrategy(autoWidthStrategy);
        return this;
    }

    public DefaultExcelBuilder fixedTitles() {
        streamExcelBuilder.fixedTitles();
        return this;
    }

    public DefaultExcelBuilder widths(int... widths) {
        streamExcelBuilder.widths(widths);
        return this;
    }

    public DefaultExcelBuilder groups(Class<?>... groups) {
        streamExcelBuilder.groups(groups);
        return this;
    }

    public Workbook build(List<?> data) {
        try {
            streamExcelBuilder.start();
            streamExcelBuilder.append(data);
            return streamExcelBuilder.build();
        } catch (Exception e) {
            try {
                streamExcelBuilder.close();
            } catch (IOException e1) {
                e1.printStackTrace();
                // do nothing
            }
            throw new RuntimeException(e);
        }
    }

    @Override
    public void close() throws IOException {
        if (streamExcelBuilder != null) {
            streamExcelBuilder.close();
        }
    }
}
