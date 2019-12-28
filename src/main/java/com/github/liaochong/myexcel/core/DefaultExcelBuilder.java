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
import java.util.Map;

/**
 * 默认excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelBuilder<T> implements Closeable {

    private DefaultStreamExcelBuilder<T> streamExcelBuilder;

    private DefaultExcelBuilder(DefaultStreamExcelBuilder<T> streamExcelBuilder) {
        streamExcelBuilder.hasStyle();
        streamExcelBuilder.widthStrategy(WidthStrategy.COMPUTE_AUTO_WIDTH);
        this.streamExcelBuilder = streamExcelBuilder;
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型ff
     *
     * @param dataType 数据的类类型
     * @return DefaultExcelBuilder
     */
    public static <T> DefaultExcelBuilder<T> of(@NonNull Class<T> dataType) {
        DefaultExcelBuilder<T> defaultExcelBuilder = new DefaultExcelBuilder<>(DefaultStreamExcelBuilder.of(dataType));
        defaultExcelBuilder.streamExcelBuilder.workbookType(WorkbookType.XLSX);
        return defaultExcelBuilder;
    }

    public static <T> DefaultExcelBuilder<T> of(@NonNull Class<T> dataType, @NonNull Workbook workbook) {
        return new DefaultExcelBuilder<>(DefaultStreamExcelBuilder.of(dataType, workbook));
    }

    /**
     * 已过时，获取实例
     *
     * @return DefaultExcelBuilder
     */
    @Deprecated
    public static DefaultExcelBuilder<Map> getInstance() {
        DefaultExcelBuilder<Map> defaultExcelBuilder = new DefaultExcelBuilder<>(DefaultStreamExcelBuilder.getInstance());
        defaultExcelBuilder.streamExcelBuilder.workbookType(WorkbookType.XLSX);
        return defaultExcelBuilder;
    }

    /**
     * 已过时，获取实例
     *
     * @param workbook workbook
     * @return DefaultExcelBuilder
     */
    @Deprecated
    public static DefaultExcelBuilder<Map> getInstance(Workbook workbook) {
        return new DefaultExcelBuilder<>(DefaultStreamExcelBuilder.getInstance(workbook));
    }

    public DefaultExcelBuilder<T> titles(@NonNull List<String> titles) {
        streamExcelBuilder.titles(titles);
        return this;
    }

    public DefaultExcelBuilder<T> sheetName(@NonNull String sheetName) {
        streamExcelBuilder.sheetName(sheetName);
        return this;
    }

    public DefaultExcelBuilder<T> fieldDisplayOrder(@NonNull List<String> fieldDisplayOrder) {
        streamExcelBuilder.fieldDisplayOrder(fieldDisplayOrder);
        return this;
    }

    public DefaultExcelBuilder<T> workbookType(@NonNull WorkbookType workbookType) {
        streamExcelBuilder.workbookType(workbookType);
        return this;
    }

    public DefaultExcelBuilder<T> noStyle() {
        streamExcelBuilder.noStyle();
        return this;
    }

    public DefaultExcelBuilder<T> widthStrategy(WidthStrategy widthStrategy) {
        streamExcelBuilder.widthStrategy(widthStrategy);
        return this;
    }

    @Deprecated
    public DefaultExcelBuilder<T> autoWidthStrategy(@NonNull AutoWidthStrategy autoWidthStrategy) {
        streamExcelBuilder.autoWidthStrategy(autoWidthStrategy);
        return this;
    }

    public DefaultExcelBuilder<T> fixedTitles() {
        streamExcelBuilder.fixedTitles();
        return this;
    }

    public DefaultExcelBuilder<T> widths(int... widths) {
        streamExcelBuilder.widths(widths);
        return this;
    }

    public DefaultExcelBuilder<T> groups(Class<?>... groups) {
        streamExcelBuilder.groups(groups);
        return this;
    }

    public Workbook build(List<T> data) {
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
