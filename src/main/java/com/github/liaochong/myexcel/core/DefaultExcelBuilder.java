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
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 默认excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
public class DefaultExcelBuilder<T> implements Closeable {

    private static final String STYLE_COMMON_TD = "border-top-style:thin;border-right-style:thin;border-bottom-style:thin;border-left-style:thin;";

    private static final String STYLE_TITLE = "font-weight:bold;font-size:14;text-align:center;vertical-align:middle;";

    private DefaultStreamExcelBuilder<T> streamExcelBuilder;

    private DefaultExcelBuilder(DefaultStreamExcelBuilder<T> streamExcelBuilder) {
        streamExcelBuilder.widthStrategy(WidthStrategy.COMPUTE_AUTO_WIDTH);
        streamExcelBuilder.style("title->" + STYLE_COMMON_TD + STYLE_TITLE, "even->" + STYLE_COMMON_TD,
                "odd->" + STYLE_COMMON_TD + "background-color:#f6f8fa;");
        this.streamExcelBuilder = streamExcelBuilder;
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @param <T>      T
     * @return DefaultExcelBuilder
     */
    public static <T> DefaultExcelBuilder<T> of(Class<T> dataType) {
        DefaultExcelBuilder<T> defaultExcelBuilder = new DefaultExcelBuilder<>(DefaultStreamExcelBuilder.of(dataType));
        defaultExcelBuilder.streamExcelBuilder.workbookType(WorkbookType.XLSX);
        return defaultExcelBuilder;
    }

    public static <T> DefaultExcelBuilder<T> of(Class<T> dataType, Workbook workbook) {
        return new DefaultExcelBuilder<>(DefaultStreamExcelBuilder.of(dataType, workbook));
    }

    /**
     * 已过时，获取实例，请使用of方法代替
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
     * 已过时，获取实例，请使用of方法代替
     *
     * @param workbook workbook
     * @return DefaultExcelBuilder
     */
    @Deprecated
    public static DefaultExcelBuilder<Map> getInstance(Workbook workbook) {
        return new DefaultExcelBuilder<>(DefaultStreamExcelBuilder.getInstance(workbook));
    }

    public DefaultExcelBuilder<T> titles(List<String> titles) {
        streamExcelBuilder.titles(titles);
        return this;
    }

    public DefaultExcelBuilder<T> sheetName(String sheetName) {
        streamExcelBuilder.sheetName(sheetName);
        return this;
    }

    public DefaultExcelBuilder<T> fieldDisplayOrder(List<String> fieldDisplayOrder) {
        streamExcelBuilder.fieldDisplayOrder(fieldDisplayOrder);
        return this;
    }

    public DefaultExcelBuilder<T> workbookType(WorkbookType workbookType) {
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
    public DefaultExcelBuilder<T> autoWidthStrategy(AutoWidthStrategy autoWidthStrategy) {
        streamExcelBuilder.autoWidthStrategy(autoWidthStrategy);
        return this;
    }

    public DefaultExcelBuilder<T> fixedTitles() {
        streamExcelBuilder.fixedTitles();
        return this;
    }

    public DefaultExcelBuilder<T> freezePane(FreezePane freezePane) {
        streamExcelBuilder.freezePane(freezePane);
        return this;
    }

    public DefaultExcelBuilder<T> widths(int... widths) {
        streamExcelBuilder.widths(widths);
        return this;
    }

    public DefaultExcelBuilder<T> width(int columnIndex, int width) {
        streamExcelBuilder.width(columnIndex, width);
        return this;
    }

    public DefaultExcelBuilder<T> hideColumns(int... columnIndexs) {
        streamExcelBuilder.hideColumns(columnIndexs);
        return this;
    }

    public DefaultExcelBuilder<T> groups(Class<?>... groups) {
        streamExcelBuilder.groups(groups);
        return this;
    }

    @Deprecated
    public DefaultExcelBuilder<T> globalStyle(String... styles) {
        return style(styles);
    }

    public DefaultExcelBuilder<T> style(String... styles) {
        streamExcelBuilder.style(styles);
        return this;
    }

    public DefaultExcelBuilder<T> titleRowHeight(int titleRowHeight) {
        streamExcelBuilder.titleRowHeight(titleRowHeight);
        return this;
    }

    public DefaultExcelBuilder<T> rowHeight(int rowHeight) {
        streamExcelBuilder.rowHeight(rowHeight);
        return this;
    }

    public DefaultExcelBuilder<T> binding(Object... applicationBeans) {
        streamExcelBuilder.binding(applicationBeans);
        return this;
    }

    public DefaultExcelBuilder<T> binding(Set<Object> applicationBeans) {
        streamExcelBuilder.binding(applicationBeans);
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
