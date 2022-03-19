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
import com.github.liaochong.myexcel.core.parser.ParseConfig;
import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.core.parser.Tr;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.core.templatehandler.TemplateHandler;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public class DefaultStreamExcelBuilder<T> extends AbstractSimpleExcelBuilder implements SimpleStreamExcelBuilder<T> {
    private static final Logger log = org.slf4j.LoggerFactory.getLogger(DefaultStreamExcelBuilder.class);
    /**
     * 设置需要渲染的数据的类类型
     */
    private final Class<T> dataType;
    /**
     * 流工厂
     */
    private HtmlToExcelStreamFactory htmlToExcelStreamFactory;
    /**
     * workbook
     */
    private Workbook workbook;
    /**
     * 待追加excel
     */
    private Path excel;
    /**
     * 任务取消
     */
    private volatile boolean cancel;
    /**
     * 分组
     */
    private Class<?>[] groups;
    /**
     * 模板处理器
     */
    private TemplateHandler templateHandler;

    private final List<CompletableFuture<Void>> asyncAppendFutures = new LinkedList<>();

    private final HtmlToExcelStreamFactory.HtmlToExcelStreamFactoryContext context = new HtmlToExcelStreamFactory.HtmlToExcelStreamFactoryContext();

    private DefaultStreamExcelBuilder(Class<T> dataType) {
        this(dataType, (Workbook) null);
    }

    private DefaultStreamExcelBuilder(Class<T> dataType, Workbook workbook) {
        super(false);
        this.dataType = dataType;
        this.workbook = workbook;
        configuration.setWidthStrategy(WidthStrategy.NO_AUTO);
        this.isMapBuild = dataType == Map.class;
    }

    private DefaultStreamExcelBuilder(Class<T> dataType, Path excel) {
        super(false);
        this.dataType = dataType;
        this.excel = excel;
        configuration.setWidthStrategy(WidthStrategy.NO_AUTO);
        this.isMapBuild = dataType == Map.class;
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @param <T>      T
     * @return DefaultStreamExcelBuilder
     */
    public static <T> DefaultStreamExcelBuilder<T> of(Class<T> dataType) {
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
    public static <T> DefaultStreamExcelBuilder<T> of(Class<T> dataType, Workbook workbook) {
        return new DefaultStreamExcelBuilder<>(dataType, workbook);
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @param excel    excel
     * @param <T>      T
     * @return DefaultStreamExcelBuilder
     */
    public static <T> DefaultStreamExcelBuilder<T> of(Class<T> dataType, Path excel) {
        return new DefaultStreamExcelBuilder<>(dataType, excel);
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType         数据的类类型
     * @param excelInputStream excelInputStream
     * @param <T>              T
     * @return DefaultStreamExcelBuilder
     */
    public static <T> DefaultStreamExcelBuilder<T> of(Class<T> dataType, InputStream excelInputStream) {
        return of(dataType, TempFileOperator.convertToFile(excelInputStream));
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

    public DefaultStreamExcelBuilder<T> titles(List<String> titles) {
        this.titles = titles;
        return this;
    }

    public DefaultStreamExcelBuilder<T> sheetName(String sheetName) {
        configuration.sheetName = sheetName;
        return this;
    }

    public DefaultStreamExcelBuilder<T> fieldDisplayOrder(List<String> fieldDisplayOrder) {
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    public DefaultStreamExcelBuilder<T> workbookType(WorkbookType workbookType) {
        if (workbook != null) {
            throw new IllegalArgumentException("Workbook type confirmed, not modifiable");
        }
        configuration.workbookType = workbookType;
        return this;
    }

    /**
     * 设置为无样式
     *
     * @return DefaultStreamExcelBuilder
     */
    public DefaultStreamExcelBuilder<T> noStyle() {
        this.styleParser.setNoStyle(true);
        return this;
    }

    public DefaultStreamExcelBuilder<T> widthStrategy(WidthStrategy widthStrategy) {
        configuration.setWidthStrategy(widthStrategy);
        return this;
    }

    @Deprecated
    public DefaultStreamExcelBuilder<T> autoWidthStrategy(AutoWidthStrategy autoWidthStrategy) {
        configuration.setWidthStrategy(AutoWidthStrategy.map(autoWidthStrategy));
        return this;
    }

    public DefaultStreamExcelBuilder<T> fixedTitles() {
        this.context.fixedTitles = true;
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
        this.context.executorService = executorService;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> hasStyle() {
        this.styleParser.setNoStyle(false);
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> capacity(int capacity) {
        if (capacity < 1) {
            throw new IllegalArgumentException("Capacity must be greater than 0");
        }
        this.context.capacity = capacity;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> pathConsumer(Consumer<Path> pathConsumer) {
        this.context.pathConsumer = pathConsumer;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder<T> groups(Class<?>... groups) {
        this.groups = groups;
        return this;
    }

    public DefaultStreamExcelBuilder<T> waitQueueSize(int waitQueueSize) {
        this.context.trWaitQueue = new LinkedBlockingQueue<>(waitQueueSize);
        return this;
    }

    @Deprecated
    public DefaultStreamExcelBuilder<T> globalStyle(String... styles) {
        return style(styles);
    }

    public DefaultStreamExcelBuilder<T> style(String... styles) {
        this.styleParser.setNoStyle(false);
        configuration.style = Arrays.stream(styles).collect(Collectors.toSet());
        return this;
    }

    public DefaultStreamExcelBuilder<T> templateHandler(Class<? extends TemplateHandler> templateHandlerClass) {
        templateHandler = ReflectUtil.newInstance(templateHandlerClass);
        return this;
    }

    public DefaultStreamExcelBuilder<T> startSheet(Consumer<Sheet> startSheetConsumer) {
        this.context.startSheetConsumer = startSheetConsumer;
        return this;
    }

    public DefaultStreamExcelBuilder<T> freezePane(FreezePane freezePane) {
        this.context.freezePane = freezePane;
        return this;
    }

    public DefaultStreamExcelBuilder<T> titleRowHeight(int titleRowHeight) {
        this.configuration.titleRowHeight = titleRowHeight;
        return this;
    }

    public DefaultStreamExcelBuilder<T> rowHeight(int rowHeight) {
        this.configuration.rowHeight = rowHeight;
        return this;
    }

    public DefaultStreamExcelBuilder<T> binding(Object... applicationBeans) {
        this.binding(new HashSet<>(Arrays.asList(applicationBeans)));
        return this;
    }

    public DefaultStreamExcelBuilder<T> binding(Set<Object> applicationBeans) {
        if (applicationBeans.isEmpty()) {
            log.warn("binding application beans failure");
            return this;
        }
        Map<Class<?>, Object> map = applicationBeans.stream().collect(Collectors.toMap(Object::getClass, b -> b));
        this.configuration.applicationBeans = map;
        return this;
    }

    public DefaultStreamExcelBuilder<T> autoMerge() {
        this.configuration.autoMerge = true;
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
        context.styleParser = styleParser;
        htmlToExcelStreamFactory = new HtmlToExcelStreamFactory(context);
        htmlToExcelStreamFactory.widthStrategy(configuration.widthStrategy);
        if (workbook == null) {
            htmlToExcelStreamFactory.workbookType(configuration.workbookType);
        }
        Table table = this.createTable();
        List<Tr> head = this.createThead();
        if (head != null) {
            htmlToExcelStreamFactory.appendTitles(head);
        }
        htmlToExcelStreamFactory.start(table, workbook);

        if (excel != null && Files.exists(excel)) {
            log.info("start reading existing excel data.");
            SaxExcelReader<T> reader = SaxExcelReader.of(dataType)
                    .readAllSheet();
            if (titleLevel > 0) {
                reader.rowFilter(row -> row.getRowNum() > titleLevel - 1);
            }
            reader.readThen(excel.toFile(), (Consumer<T>) this::append);
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
        if (isMapBuild) {
            List<Pair<? extends Class, ?>> contents = assemblingMapContents((Map<String, Object>) data);
            this.appendTr(contents);
            return;
        }
        if (hasMultiColumn) {
            List<List<Pair<? extends Class, ?>>> contents = this.getMultiRenderContent(data, filteredFields);
            for (List<Pair<? extends Class, ?>> content : contents) {
                this.appendTr(content);
            }
        } else {
            List<Pair<? extends Class, ?>> contents = this.getOriginalRenderContent(data, filteredFields);
            this.appendTr(contents);
        }
    }

    private void appendTr(List<Pair<? extends Class, ?>> contents) {
        Tr tr = this.createTr(contents);
        htmlToExcelStreamFactory.append(tr);
    }

    public <E> void append(String templateFilePath, Map<String, E> renderData) {
        templateHandler.classpathTemplate(templateFilePath);
        this.doAppend(renderData);
    }

    public <E> void append(String templateDir, String templateFileName, Map<String, E> renderData) {
        templateHandler.fileTemplate(templateDir, templateFileName);
        this.doAppend(renderData);
    }

    public void asyncAppend(ListSupplier<T> supplier) {
        CompletableFuture<Void> future;
        if (this.context.executorService == null) {
            future = CompletableFuture.runAsync(() -> {
                this.append(supplier.getAsList());
            });
        } else {
            future = CompletableFuture.runAsync(() -> {
                this.append(supplier.getAsList());
            }, this.context.executorService);
        }
        synchronized (this) {
            asyncAppendFutures.add(future);
        }
    }

    public void asyncAppend(Supplier<T> supplier) {
        CompletableFuture<Void> future;
        if (this.context.executorService == null) {
            future = CompletableFuture.runAsync(() -> {
                this.append(supplier.get());
            });
        } else {
            future = CompletableFuture.runAsync(() -> {
                this.append(supplier.get());
            }, this.context.executorService);
        }
        synchronized (this) {
            asyncAppendFutures.add(future);
        }
    }

    @Override
    public Workbook build() {
        joinAsyncAppendFutures();
        return htmlToExcelStreamFactory.build();
    }

    @Override
    public List<Path> buildAsPaths() {
        joinAsyncAppendFutures();
        return htmlToExcelStreamFactory.buildAsPaths();
    }

    @Override
    public Path buildAsZip(String fileName) {
        joinAsyncAppendFutures();
        return htmlToExcelStreamFactory.buildAsZip(fileName);
    }

    @Override
    public void close() throws IOException {
        if (htmlToExcelStreamFactory != null) {
            htmlToExcelStreamFactory.clear();
            htmlToExcelStreamFactory = null;
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

    private <E> void doAppend(Map<String, E> renderData) {
        List<Table> tables;
        try {
            tables = templateHandler.render(renderData, new ParseConfig(configuration.widthStrategy));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        if (tables == null || tables.isEmpty()) {
            return;
        }
        for (Table table : tables) {
            table.trList.forEach(htmlToExcelStreamFactory::append);
        }
    }

    private void joinAsyncAppendFutures() {
        if (!asyncAppendFutures.isEmpty()) {
            asyncAppendFutures.forEach(CompletableFuture::join);
        }
    }

    protected Table createTable() {
       return createTable();
   }
}
