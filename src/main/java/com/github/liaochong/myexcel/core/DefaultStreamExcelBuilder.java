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
import java.lang.reflect.Field;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultStreamExcelBuilder extends AbstractSimpleExcelBuilder implements SimpleStreamExcelBuilder {
    /**
     * 已排序字段
     */
    private List<Field> filteredFields;
    /**
     * 设置需要渲染的数据的类类型
     */
    private Class<?> dataType;
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
    private boolean cancel;
    /**
     * 分组
     */
    private Class<?>[] groups;
    /**
     * 等待队列
     */
    private int waitQueueSize = Runtime.getRuntime().availableProcessors() * 2;

    private DefaultStreamExcelBuilder() {
        noStyle = true;
        widthStrategy = WidthStrategy.NO_AUTO;
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

    public static DefaultStreamExcelBuilder getInstance() {
        return new DefaultStreamExcelBuilder();
    }

    public DefaultStreamExcelBuilder titles(@NonNull List<String> titles) {
        this.titles = titles;
        return this;
    }

    public DefaultStreamExcelBuilder sheetName(@NonNull String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public DefaultStreamExcelBuilder fieldDisplayOrder(@NonNull List<String> fieldDisplayOrder) {
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    /**
     * SXSSF模式下设置窗口保留行数，采用默认值是最佳实践，不建议修改
     *
     * @param rowAccessWindowSize 内存保留行数
     * @return DefaultStreamExcelBuilder
     */
    @Deprecated
    public DefaultStreamExcelBuilder rowAccessWindowSize(int rowAccessWindowSize) {
        // do nothing
        return this;
    }

    public DefaultStreamExcelBuilder workbookType(@NonNull WorkbookType workbookType) {
        this.workbookType = workbookType;
        return this;
    }

    /**
     * 设置为无样式
     *
     * @return DefaultStreamExcelBuilder
     */
    public DefaultStreamExcelBuilder noStyle() {
        this.noStyle = true;
        return this;
    }

    public DefaultStreamExcelBuilder widthStrategy(@NonNull WidthStrategy widthStrategy) {
        this.widthStrategy = widthStrategy;
        return this;
    }

    @Deprecated
    public DefaultStreamExcelBuilder autoWidthStrategy(@NonNull AutoWidthStrategy autoWidthStrategy) {
        this.widthStrategy = AutoWidthStrategy.map(autoWidthStrategy);
        return this;
    }

    public DefaultStreamExcelBuilder fixedTitles() {
        this.fixedTitles = true;
        return this;
    }

    public DefaultStreamExcelBuilder widths(int... widths) {
        if (widths.length == 0) {
            return this;
        }
        this.widths = new HashMap<>(widths.length);
        for (int i = 0, size = widths.length; i < size; i++) {
            this.widths.put(i, widths[i]);
        }
        customWidthMap = this.widths;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder threadPool(ExecutorService executorService) {
        this.executorService = executorService;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder hasStyle() {
        this.noStyle = false;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder capacity(int capacity) {
        if (capacity < 1) {
            throw new IllegalArgumentException("Capacity must be greater than 0");
        }
        this.capacity = capacity;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder pathConsumer(Consumer<Path> pathConsumer) {
        this.pathConsumer = pathConsumer;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder groups(Class<?>... groups) {
        this.groups = groups;
        return this;
    }

    public DefaultStreamExcelBuilder waitQueueSize(int waitQueueSize) {
        this.waitQueueSize = waitQueueSize;
        return this;
    }

    /**
     * 流式构建启动，包含一些初始化操作
     *
     * @return DefaultExcelBuilder
     */
    @Override
    public DefaultStreamExcelBuilder start() {
        htmlToExcelStreamFactory = new HtmlToExcelStreamFactory(waitQueueSize, executorService, pathConsumer, capacity, fixedTitles);
        htmlToExcelStreamFactory.workbookType(workbookType).widthStrategy(widthStrategy);

        if (dataType != null) {
            ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
            filteredFields = getFilteredFields(classFieldContainer, groups);
        }

        this.initStyleMap();
        Table table = this.createTable();
        htmlToExcelStreamFactory.start(table, workbook);

        List<Tr> head = this.createThead();
        if (head != null) {
            htmlToExcelStreamFactory.appendTitles(head);
        }
        return this;
    }

    @SuppressWarnings("unchecked")
    @Override
    public void append(List<?> data) {
        if (cancel) {
            log.info("Canceled build task");
            return;
        }
        if (data == null || data.isEmpty()) {
            return;
        }
        boolean isMapBuild = data.stream().anyMatch(d -> d instanceof Map);
        if (isMapBuild) {
            for (Object datum : data) {
                Map<String, Object> d = (Map<String, Object>) datum;
                List<Pair<? extends Class, ?>> contents = new ArrayList<>(d.size());
                for (String fieldName : fieldDisplayOrder) {
                    Object val = d.get(fieldName);
                    contents.add(Pair.of(val == null ? String.class : val.getClass(), val));
                }
                Tr tr = this.createTr(contents);
                htmlToExcelStreamFactory.append(tr);
            }
        } else {
            for (Object datum : data) {
                List<Pair<? extends Class, ?>> contents = getRenderContent(datum, filteredFields);
                Tr tr = this.createTr(contents);
                htmlToExcelStreamFactory.append(tr);
            }
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    public <T> void append(T data) {
        if (cancel) {
            log.info("Canceled build task");
            return;
        }
        if (data == null) {
            return;
        }
        List<Pair<? extends Class, ?>> contents;
        if (data instanceof Map) {
            Map<String, Object> d = (Map<String, Object>) data;
            contents = new ArrayList<>(d.size());
            for (String fieldName : fieldDisplayOrder) {
                Object val = d.get(fieldName);
                contents.add(Pair.of(val == null ? String.class : val.getClass(), val));
            }
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

    public void cancle() {
        cancel = true;
        htmlToExcelStreamFactory.clear();
    }

    public void clear() {
        htmlToExcelStreamFactory.clear();
    }
}
