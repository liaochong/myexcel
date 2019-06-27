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

import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.core.parser.Tr;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * HtmlToExcelStreamFactory 流工厂
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
class HtmlToExcelStreamFactory extends AbstractExcelFactory {

    private static final int XLSX_MAX_ROW_COUNT = 1048576;

    private static final int XLS_MAX_ROW_COUNT = 65536;

    static final int DEFAULT_WAIT_SIZE = Runtime.getRuntime().availableProcessors();

    private static final List<Tr> STOP_FLAG_LIST = new ArrayList<>();

    private int maxRowCountOfSheet = XLSX_MAX_ROW_COUNT;

    private Sheet sheet;

    private BlockingQueue<List<Tr>> trWaitQueue;

    private boolean stop;

    private boolean exception;

    private long startTime;

    private String sheetName = "Sheet";

    private Map<Integer, Integer> colWidthMap;

    private int rowNum;

    private int sheetNum;

    private int maxColIndex;

    /**
     * 文件分割,excel容量
     */
    private int capacity;

    /**
     * 计数器
     */
    private int count;

    private List<Tr> titles;

    private List<Path> paths;

    private List<CompletableFuture> futures;

    private Consumer<Path> pathConsumer;

    /**
     * 线程池
     */
    private ExecutorService executorService;

    public HtmlToExcelStreamFactory(int waitSize, ExecutorService executorService,
                                    Consumer<Path> pathConsumer, int capacity) {
        this.trWaitQueue = new ArrayBlockingQueue<>(waitSize);
        this.executorService = executorService;
        this.pathConsumer = pathConsumer;
        this.capacity = capacity;
    }

    public void start(Table table, Workbook workbook) {
        log.info("Start streaming building excel");
        if (Objects.nonNull(workbook)) {
            this.workbook = workbook;
        }
        startTime = System.currentTimeMillis();

        if (Objects.isNull(this.workbook)) {
            workbookType(WorkbookType.SXLSX);
        }
        if (workbook instanceof HSSFWorkbook) {
            maxRowCountOfSheet = XLS_MAX_ROW_COUNT;
        }
        initCellStyle(workbook);
        if (Objects.nonNull(table)) {
            sheetName = Objects.isNull(table.getCaption()) || table.getCaption().length() < 1 ? sheetName : table.getCaption();
        }
        this.sheet = this.workbook.createSheet(sheetName);
        paths = new ArrayList<>();
        if (Objects.isNull(executorService)) {
            Thread thread = new Thread(this::receive);
            thread.setName("Excel-builder-1");
            thread.start();
        } else {
            futures = new ArrayList<>();
            CompletableFuture.runAsync(this::receive, executorService);
        }
    }

    public void appendTitles(List<Tr> trList) {
        this.titles = trList;
        this.append(trList);
    }

    public void append(List<Tr> trList) {
        if (exception) {
            log.error("Received a termination command,an exception occurred while processing");
            throw new UnsupportedOperationException("Received a termination command");
        }
        if (stop) {
            log.error("Received a termination command,the build method has been called");
            throw new UnsupportedOperationException("Received a termination command");
        }
        if (Objects.isNull(trList) || trList.isEmpty()) {
            log.warn("This list is empty and will be discarded");
            return;
        }
        try {
            trWaitQueue.put(trList);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
    }

    private void receive() {
        List<Tr> trList = this.getTrListFromQueue();
        if (maxColIndex == 0 && !trList.isEmpty()) {
            int tdSize = trList.get(0).getTdList().size();
            maxColIndex = tdSize > 0 ? tdSize - 1 : 0;
        }
        int appendSize = 0;
        try {
            while (trList != STOP_FLAG_LIST) {
                log.info("Received data size:{},current waiting queue size:{}", trList.size(), trWaitQueue.size());
                for (Tr tr : trList) {
                    if (capacity > 0 && count == capacity) {
                        // 上一份数据保存
                        this.storeToTempFile();
                        // 开启下一份数据
                        this.initNewWorkbook();
                    }
                    if (rowNum == maxRowCountOfSheet) {
                        sheetNum++;
                        this.setColWidth(colWidthMap, sheet, maxColIndex);
                        colWidthMap = null;
                        sheet = workbook.createSheet(sheetName + " (" + sheetNum + ")");
                        rowNum = 0;
                    }
                    tr.setIndex(rowNum);
                    tr.getTdList().forEach(td -> {
                        td.setRow(rowNum);
                    });
                    rowNum++;
                    count++;
                    this.createRow(tr, sheet);
                }
                appendSize++;
                Map<Integer, Integer> colWidthMap = this.getColMaxWidthMap(trList);
                if (Objects.isNull(this.colWidthMap)) {
                    this.colWidthMap = new HashMap<>(colWidthMap.size());
                }
                colWidthMap.forEach((k, v) -> {
                    Integer val = this.colWidthMap.get(k);
                    if (Objects.isNull(val) || v > val) {
                        this.colWidthMap.put(k, v);
                    }
                });
                trList = this.getTrListFromQueue();
            }
            log.info("End of reception,append size:{}", appendSize);
        } catch (Exception e) {
            log.error("An exception occurred while processing", e);
            exception = true;
            try {
                workbook.close();
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            trWaitQueue.clear();
            trWaitQueue = null;
            TempFileOperator.deleteTempFiles(paths);
        }
    }

    private List<Tr> getTrListFromQueue() {
        try {
            return trWaitQueue.take();
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public Workbook build() {
        if (exception) {
            throw new IllegalStateException("An exception occurred while processing");
        }
        this.stop = true;
        while (!trWaitQueue.isEmpty()) {
            // wait all tr received
        }
        try {
            trWaitQueue.put(STOP_FLAG_LIST);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        while (!trWaitQueue.isEmpty()) {
            // wait all tr received
        }
        this.setColWidth(colWidthMap, sheet, maxColIndex);
        this.freezePane(0, sheet);
        log.info("Build Excel success,takes {} ms", System.currentTimeMillis() - startTime);
        return workbook;
    }

    public List<Path> buildAsPaths() {
        if (exception) {
            throw new IllegalStateException("An exception occurred while processing");
        }
        this.stop = true;
        while (!trWaitQueue.isEmpty()) {
            // wait all tr received
        }
        try {
            trWaitQueue.put(STOP_FLAG_LIST);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        while (!trWaitQueue.isEmpty()) {
            // wait all tr received
        }
        this.storeToTempFile();
        if (Objects.nonNull(futures)) {
            futures.forEach(CompletableFuture::join);
        }
        log.info("Build Excel success,takes {} ms", System.currentTimeMillis() - startTime);
        return paths.stream().filter(path -> Objects.nonNull(path) && path.toFile().exists()).collect(Collectors.toList());
    }

    private void storeToTempFile() {
        boolean isXls = workbook instanceof HSSFWorkbook;
        String suffix = isXls ? Constants.XLS : Constants.XLSX;
        Path path = TempFileOperator.createTempFile("s_t_r_p", suffix);
        paths.add(path);
        try {
            if (Objects.nonNull(executorService)) {
                Workbook tempWorkbook = workbook;
                Sheet tempSheet = sheet;
                Map<Integer, Integer> tempColWidthMap = colWidthMap;
                CompletableFuture future = CompletableFuture.runAsync(() -> {
                    this.setColWidth(tempColWidthMap, tempSheet, maxColIndex);
                    this.freezePane(0, tempSheet);
                    try {
                        FileExportUtil.export(tempWorkbook, path.toFile());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                    if (Objects.nonNull(pathConsumer)) {
                        pathConsumer.accept(path);
                    }
                }, executorService);
                futures.add(future);
            } else {
                this.setColWidth(colWidthMap, sheet, maxColIndex);
                this.freezePane(0, sheet);
                FileExportUtil.export(workbook, path.toFile());
                if (Objects.nonNull(pathConsumer)) {
                    pathConsumer.accept(path);
                }
            }
        } catch (IOException e) {
            TempFileOperator.deleteTempFiles(paths);
            throw new RuntimeException(e);
        }
    }

    private void initNewWorkbook() {
        boolean isXls = workbook instanceof HSSFWorkbook;
        workbook = null;
        workbookType(isXls ? WorkbookType.XLS : WorkbookType.SXLSX);
        sheetNum = 0;
        rowNum = 0;
        count = 0;
        colWidthMap = null;
        clearCache();
        initCellStyle(workbook);
        sheet = workbook.createSheet(sheetName);
        // 标题构建
        if (Objects.isNull(titles)) {
            return;
        }
        for (Tr titleTr : titles) {
            titleTr.setIndex(rowNum);
            titleTr.getTdList().forEach(td -> {
                td.setRow(rowNum);
            });
            rowNum++;
            count++;
            this.createRow(titleTr, sheet);
        }
    }
}
