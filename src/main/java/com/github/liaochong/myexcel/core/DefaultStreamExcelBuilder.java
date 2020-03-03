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
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultStreamExcelBuilder<T> extends AbstractSimpleExcelBuilder implements SimpleStreamExcelBuilder<T> {
    /**
     * 设置需要渲染的数据的类类型
     */
    private Class<T> dataType;
    /**
     * 是否固定标题
     */
    private boolean fixedTitles;
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
    /**
     * 文件分割,excel容量
     */
    private int capacity;
    /**
     * path消费者
     */
    private Consumer<Path> pathConsumer;
    /**
     * 任务取消
     */
    private volatile boolean cancel;
    /**
     * 分组
     */
    private Class<?>[] groups;
    /**
     * 等待队列
     */
    private int waitQueueSize = Runtime.getRuntime().availableProcessors() * 2;

    private DefaultStreamExcelBuilder(Class<T> dataType) {
        this(dataType, null);
    }

    private DefaultStreamExcelBuilder(Class<T> dataType, Workbook workbook) {
        super(false);
        this.dataType = dataType;
        this.workbook = workbook;
        globalSetting.setWidthStrategy(WidthStrategy.NO_AUTO);
        this.isMapBuild = dataType == Map.class;
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @param <T>      T
     * @return DefaultStreamExcelBuilder
     */
    public static <T> DefaultStreamExcelBuilder<T> of(@NonNull Class<T> dataType) {
        return new DefaultStreamExcelBuilder<>(dataType);
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @param workbook workbook
     * @param <T>      T
     * @return DefaultStreamExcelBuilder
     */
    public static <T> DefaultStreamExcelBuilder<T> of(@NonNull Class<T> dataType, @NonNull Workbook workbook) {
        return new DefaultStreamExcelBuilder<>(dataType, workbook);
    }

    /**
     * 已过时，请使用of方法代替
     * 4.0版本移除
     *
     * @return DefaultStreamExcelBuilder
     */
    @Deprecated
    public static DefaultStreamExcelBuilder<Map> getInstance() {
        return new DefaultStreamExcelBuilder<>(Map.class);
    }

    /**
     * 已过时，请使用of方法代替
     * 4.0版本移除
     *
     * @param workbook 工作簿
     * @return DefaultStreamExcelBuilder
     */
    @Deprecated
    public static DefaultStreamExcelBuilder<Map> getInstance(Workbook workbook) {
        return new DefaultStreamExcelBuilder<>(Map.class, workbook);
    }

    public DefaultStreamExcelBuilder<T> titles(@NonNull List<String> titles) {
        this.titles = titles;
        return this;
    }

    public DefaultStreamExcelBuilder<T> sheetName(@NonNull String sheetName) {
        globalSetting.setSheetName(sheetName);
        return this;
    }

    public DefaultStreamExcelBuilder<T> fieldDisplayOrder(@NonNull List<String> fieldDisplayOrder) {
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    public DefaultStreamExcelBuilder<T> workbookType(@NonNull WorkbookType workbookType) {
        if (workbook != null) {
            throw new IllegalArgumentException("Workbook type confirmed, not modifiable");
        }
        globalSetting.setWorkbookType(workbookType);
        return this;
    }

    /**
     * 设置为无样式
     *
     * @return DefaultStreamExcelBuilder
     */
    public DefaultStreamExcelBuilder<T> noStyle() {
        this.noStyle = true;
        return this;
    }

    public DefaultStreamExcelBuilder<T> widthStrategy(@NonNull WidthStrategy widthStrategy) {
        globalSetting.setWidthStrategy(widthStrategy);
        return this;
    }

    @Deprecated
    public DefaultStreamExcelBuilder<T> autoWidthStrategy(@NonNull AutoWidthStrategy autoWidthStrategy) {
        globalSetting.setWidthStrategy(AutoWidthStrategy.map(autoWidthStrategy));
        return this;
    }

    public DefaultStreamExcelBuilder<T> fixedTitles() {
        this.fixedTitles = true;
        return this;
    }

    public DefaultStreamExcelBuilder<T> widths(int... widths) {
        for (int i = 0; i < widths.length; i++) {
            customWidthMap.put(i, widths[i]);
        }
        return this;
    }

    public DefaultStreamExcelBuilder<T> width(int columnIndex, int width) {
        customWidthMap.put(columnIndex, width);
        return this;
    }

    public DefaultStreamExcelBuilder<T> hideColumns(int... columnIndexs) {
        for (int columnIndex : columnIndexs) {
            width(columnIndex, 0);
        }
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> threadPool(ExecutorService executorService) {
        this.executorService = executorService;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> hasStyle() {
        this.noStyle = false;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> capacity(int capacity) {
        if (capacity < 1) {
            throw new IllegalArgumentException("Capacity must be greater than 0");
        }
        this.capacity = capacity;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> pathConsumer(Consumer<Path> pathConsumer) {
        this.pathConsumer = pathConsumer;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> groups(Class<?>... groups) {
        this.groups = groups;
        return this;
    }

    public DefaultStreamExcelBuilder<T> waitQueueSize(int waitQueueSize) {
        this.waitQueueSize = waitQueueSize;
        return this;
    }

    @Deprecated
    public DefaultStreamExcelBuilder<T> globalStyle(String... styles) {
        return style(styles);
    }

    public DefaultStreamExcelBuilder<T> style(String... styles) {
        this.noStyle = false;
        globalSetting.setStyle(Arrays.stream(styles).collect(Collectors.toSet()));
        return this;
    }

    /**
     * 流式构建启动，包含一些初始化操作
     *
     * @return DefaultExcelBuilder
     */
    @Override
    public DefaultStreamExcelBuilder<T> start() {
        if (isMapBuild) {
            this.parseGlobalStyle();
        } else {
            ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
            filteredFields = getFilteredFields(classFieldContainer, groups);
        }
        htmlToExcelStreamFactory = new HtmlToExcelStreamFactory(waitQueueSize, executorService, pathConsumer, capacity, fixedTitles);
        htmlToExcelStreamFactory.widthStrategy(globalSetting.getWidthStrategy());
        if (workbook == null) {
            htmlToExcelStreamFactory.workbookType(globalSetting.getWorkbookType());
        }
        Table table = this.createTable();
        htmlToExcelStreamFactory.start(table, workbook);

        List<Tr> head = this.createThead();
        if (head != null) {
            htmlToExcelStreamFactory.appendTitles(head);
        }
        return this;
    }

    @Override
    public void append(List<T> dataList) {
        if (cancel) {
            log.info("Canceled build task");
            return;
        }
        if (dataList == null || dataList.isEmpty()) {
            return;
        }
        for (T data : dataList) {
            this.append(data);
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    public void append(T data) {
        if (cancel) {
            log.info("Canceled build task");
            return;
        }
        if (data == null) {
            return;
        }
        List<Pair<? extends Class, ?>> contents;
        if (isMapBuild) {
            contents = assemblingMapContents((Map<String, Object>) data);
        } else {
            contents = getRenderContent(data, filteredFields);
        }
        Tr tr = this.createTr(contents);
        htmlToExcelStreamFactory.append(tr);
    }

    @Override
    public Workbook build() {
        return htmlToExcelStreamFactory.build();
    }

    @Override
    public List<Path> buildAsPaths() {
        return htmlToExcelStreamFactory.buildAsPaths();
    }

    @Override
    public Path buildAsZip(String fileName) {
        return htmlToExcelStreamFactory.buildAsZip(fileName);
    }

    @Override
    public void close() throws IOException {
        if (htmlToExcelStreamFactory != null) {
            htmlToExcelStreamFactory.closeWorkbook();
        }
    }

    public void cancel() {
        cancel = true;
        htmlToExcelStreamFactory.cancel();
    }

    /**
     * clear方法仅可在异常情况下调用
     */
    public void clear() {
        htmlToExcelStreamFactory.clear();
    }
}
