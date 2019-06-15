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

import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.core.parser.Tr;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.ExecutorService;

/**
 * @author liaochong
 * @version 1.0
 */
public class DefaultStreamExcelBuilder extends AbstractSimpleExcelBuilder implements SimpleStreamExcelBuilder {
    /**
     * 线程池
     */
    private ExecutorService executorService;
    /**
     * 流工厂
     */
    private HtmlToExcelStreamFactory htmlToExcelStreamFactory;
    /**
     * workbook
     */
    private Workbook workbook;

    private DefaultStreamExcelBuilder() {
        noStyle = true;
        autoWidthStrategy = AutoWidthStrategy.NO_AUTO;
        workbookType = WorkbookType.SXLSX;
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @return DefaultExcelBuilder
     */
    public static DefaultStreamExcelBuilder of(@NonNull Class<?> dataType) {
        DefaultStreamExcelBuilder defaultStreamExcelBuilder = new DefaultStreamExcelBuilder();
        defaultStreamExcelBuilder.dataType = dataType;
        return defaultStreamExcelBuilder;
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @param workbook workbook
     * @return DefaultExcelBuilder
     */
    public static DefaultStreamExcelBuilder of(@NonNull Class<?> dataType, @NonNull Workbook workbook) {
        DefaultStreamExcelBuilder defaultStreamExcelBuilder = new DefaultStreamExcelBuilder();
        defaultStreamExcelBuilder.dataType = dataType;
        defaultStreamExcelBuilder.workbook = workbook;
        return defaultStreamExcelBuilder;
    }

    @Override
    public DefaultStreamExcelBuilder rowAccessWindowSize(int rowAccessWindowSize) {
        super.rowAccessWindowSize(rowAccessWindowSize);
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder workbookType(@NonNull WorkbookType workbookType) {
        super.workbookType(workbookType);
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder threadPool(@NonNull ExecutorService executorService) {
        this.executorService = executorService;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder sheetName(@NonNull String sheetName) {
        super.sheetName(sheetName);
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder hasStyle() {
        this.noStyle = false;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder autoWidthStrategy(@NonNull AutoWidthStrategy autoWidthStrategy) {
        super.autoWidthStrategy(autoWidthStrategy);
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder fixedTitles() {
        this.fixedTitles = true;
        return this;
    }

    /**
     * 流式构建启动，包含一些初始化操作，等待队列容量采用CPU核心数目
     *
     * @param groups 分组
     * @return DefaultExcelBuilder
     */
    public DefaultStreamExcelBuilder start(Class<?>... groups) {
        this.start(HtmlToExcelStreamFactory.DEFAULT_WAIT_SIZE, groups);
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder start(int waitQueueSize, Class<?>... groups) {
        Objects.requireNonNull(dataType);
        htmlToExcelStreamFactory = new HtmlToExcelStreamFactory(waitQueueSize, executorService);
        htmlToExcelStreamFactory.rowAccessWindowSize(rowAccessWindowSize).workbookType(workbookType).autoWidthStrategy(autoWidthStrategy);

        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
        filteredFields = getFilteredFields(classFieldContainer, groups);

        this.initStyleMap();
        Table table = this.createTable();
        htmlToExcelStreamFactory.start(table, workbook);

        Tr head = this.createThead();
        if (Objects.isNull(head)) {
            return this;
        }
        List<Tr> headList = new ArrayList<>();
        headList.add(head);
        htmlToExcelStreamFactory.append(headList);
        return this;
    }

    @Override
    public void append(List<?> data) {
        if (Objects.isNull(data) || data.isEmpty()) {
            return;
        }
        List<List<Pair<Class, Object>>> contents = getRenderContent(data, filteredFields);
        List<Tr> trList = this.createTbody(contents, 0);
        htmlToExcelStreamFactory.append(trList);
    }

    @Override
    public Workbook build() {
        if (fixedTitles && Objects.nonNull(titles) && !titles.isEmpty()) {
            FreezePane freezePane = new FreezePane(1, titles.size());
            htmlToExcelStreamFactory.freezePanes(freezePane);
        }
        return htmlToExcelStreamFactory.build();
    }

    @Override
    public Workbook build(List<?> data, Class<?>... groups) {
        throw new UnsupportedOperationException();
    }
}
