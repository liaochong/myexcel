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
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.function.Consumer;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

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

    private static final Tr STOP_FLAG = new Tr(-1, 0);

    private int maxRowCountOfSheet = XLSX_MAX_ROW_COUNT;

    private Sheet sheet;

    private BlockingQueue<Tr> trWaitQueue;

    private boolean stop;

    private boolean exception;

    private long startTime;

    private String sheetName = "Sheet";

    private Map<Integer, Integer> colWidthMap = new HashMap<>();

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
    /**
     * 是否固定标题
     */
    private boolean fixedTitles;

    public HtmlToExcelStreamFactory(int waitSize, ExecutorService executorService,
                                    Consumer<Path> pathConsumer,
                                    int capacity,
                                    boolean fixedTitles) {
        this.trWaitQueue = new LinkedBlockingQueue<>(waitSize);
        this.executorService = executorService;
        this.pathConsumer = pathConsumer;
        this.capacity = capacity;
        this.fixedTitles = fixedTitles;
    }

    public void start(Table table, Workbook workbook) {
        log.info("Start build excel");
        if (workbook != null) {
            this.workbook = workbook;
        }
        startTime = System.currentTimeMillis();

        if (this.workbook == null) {
            workbookType(WorkbookType.SXLSX);
        }
        if (workbook instanceof HSSFWorkbook) {
            maxRowCountOfSheet = XLS_MAX_ROW_COUNT;
        }
        initCellStyle(workbook);
        if (table != null) {
            sheetName = Objects.isNull(table.getCaption()) || table.getCaption().length() < 1 ? sheetName : table.getCaption();
        }
        this.sheet = this.workbook.createSheet(sheetName);
        paths = new ArrayList<>();
        if (executorService == null) {
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
        trList.forEach(this::append);
    }

    public void append(Tr tr) {
        if (exception) {
            log.error("Received a termination command,an exception occurred while processing");
            throw new UnsupportedOperationException("Received a termination command");
        }
        if (stop) {
            log.error("Received a termination command,the build method has been called");
            throw new UnsupportedOperationException("Received a termination command");
        }
        if (tr == null) {
            log.warn("This tr is null and will be discarded");
            return;
        }
        try {
            trWaitQueue.put(tr);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
    }

    private void receive() {
        Tr tr = this.getTrFromQueue();
        if (maxColIndex == 0) {
            int tdSize = tr.getTdList().size();
            maxColIndex = tdSize > 0 ? tdSize - 1 : 0;
        }
        int totalSize = 0;
        try {
            while (tr != STOP_FLAG) {
                if (capacity > 0 && count == capacity) {
                    // 上一份数据保存
                    this.storeToTempFile();
                    // 开启下一份数据
                    this.initNewWorkbook();
                }
                if (rowNum == maxRowCountOfSheet) {
                    sheetNum++;
                    this.setColWidth(colWidthMap, sheet, maxColIndex);
                    colWidthMap = new HashMap<>();
                    sheet = workbook.createSheet(sheetName + " (" + sheetNum + ")");
                    rowNum = 0;
                    this.setTitles();
                }
                appendRow(tr);
                totalSize++;
                tr.getColWidthMap().forEach((k, v) -> {
                    Integer val = this.colWidthMap.get(k);
                    if (val == null || v > val) {
                        this.colWidthMap.put(k, v);
                    }
                });
                tr = this.getTrFromQueue();
            }
            log.info("Total size:{}", totalSize);
        } catch (Exception e) {
            log.error("An exception occurred while processing", e);
            exception = true;
            closeWorkbook();
            trWaitQueue.clear();
            trWaitQueue = null;
            TempFileOperator.deleteTempFiles(paths);
        }
    }

    private Tr getTrFromQueue() {
        try {
            return trWaitQueue.take();
        } catch (InterruptedException e) {
            closeWorkbook();
            throw new RuntimeException(e);
        }
    }

    @Override
    public Workbook build() {
        waiting();
        this.setColWidth(colWidthMap, sheet, maxColIndex);
        this.freezeTitles(workbook);
        log.info("Build Excel success,takes {} ms", System.currentTimeMillis() - startTime);
        return workbook;
    }

    List<Path> buildAsPaths() {
        waiting();
        this.storeToTempFile();
        if (futures != null) {
            futures.forEach(CompletableFuture::join);
        }
        log.info("Build Excel success,takes {} ms", System.currentTimeMillis() - startTime);
        return paths.stream().filter(path -> Objects.nonNull(path) && path.toFile().exists()).collect(Collectors.toList());
    }

    protected void waiting() {
        if (exception) {
            throw new IllegalStateException("An exception occurred while processing");
        }
        this.stop = true;
        while (!trWaitQueue.isEmpty()) {
            // wait all tr received
        }
        try {
            trWaitQueue.put(STOP_FLAG);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        while (!trWaitQueue.isEmpty()) {
            // wait all tr received
        }
    }

    private void storeToTempFile() {
        boolean isXls = workbook instanceof HSSFWorkbook;
        String suffix = isXls ? Constants.XLS : Constants.XLSX;
        Path path = TempFileOperator.createTempFile("s_t_r_p", suffix);
        paths.add(path);
        try {
            if (executorService != null) {
                Workbook tempWorkbook = workbook;
                Sheet tempSheet = sheet;
                Map<Integer, Integer> tempColWidthMap = colWidthMap;
                CompletableFuture future = CompletableFuture.runAsync(() -> {
                    this.setColWidth(tempColWidthMap, tempSheet, maxColIndex);
                    this.freezeTitles(tempWorkbook);
                    try {
                        FileExportUtil.export(tempWorkbook, path.toFile());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                    if (pathConsumer != null) {
                        pathConsumer.accept(path);
                    }
                }, executorService);
                futures.add(future);
            } else {
                this.setColWidth(colWidthMap, sheet, maxColIndex);
                this.freezeTitles(workbook);
                FileExportUtil.export(workbook, path.toFile());
                if (Objects.nonNull(pathConsumer)) {
                    pathConsumer.accept(path);
                }
            }
        } catch (IOException e) {
            closeWorkbook();
            TempFileOperator.deleteTempFiles(paths);
            throw new RuntimeException(e);
        }
    }

    private void freezeTitles(Workbook workbook) {
        if (fixedTitles && titles != null) {
            for (int i = 0, size = workbook.getNumberOfSheets(); i < size; i++) {
                workbook.getSheetAt(i).createFreezePane(0, titles.size());
            }
        }
    }

    private void initNewWorkbook() {
        boolean isXls = workbook instanceof HSSFWorkbook;
        workbook = null;
        workbookType(isXls ? WorkbookType.XLS : WorkbookType.SXLSX);
        sheetNum = 0;
        rowNum = 0;
        count = 0;
        colWidthMap = new HashMap<>();
        clearCache();
        initCellStyle(workbook);
        sheet = workbook.createSheet(sheetName);
        // 标题构建
        if (titles == null) {
            return;
        }
        this.setTitles();
    }

    private void setTitles() {
        for (Tr titleTr : titles) {
            appendRow(titleTr);
        }
    }

    private void appendRow(Tr tr) {
        tr.setIndex(rowNum);
        tr.getTdList().forEach(td -> {
            td.setRow(rowNum);
        });
        rowNum++;
        count++;
        this.createRow(tr, sheet);
    }

    Path buildAsZip(String fileName) {
        Objects.requireNonNull(fileName);
        waiting();
        boolean isXls = workbook instanceof HSSFWorkbook;
        this.storeToTempFile();
        if (Objects.nonNull(futures)) {
            futures.forEach(CompletableFuture::join);
        }
        String suffix = isXls ? Constants.XLS : Constants.XLSX;
        Path zipFile = TempFileOperator.createTempFile(fileName, ".zip");
        try (ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(zipFile))) {
            for (int i = 1, size = paths.size(); i <= size; i++) {
                Path path = paths.get(i - 1);
                ZipEntry zipEntry = new ZipEntry(fileName + " (" + i + ")" + suffix);
                out.putNextEntry(zipEntry);
                out.write(Files.readAllBytes(path));
                out.closeEntry();
            }
        } catch (IOException e) {
            closeWorkbook();
            throw new RuntimeException(e);
        } finally {
            TempFileOperator.deleteTempFiles(paths);
        }
        return zipFile;
    }

    public void clean() {
        waiting();
        closeWorkbook();
        TempFileOperator.deleteTempFiles(paths);
    }

}
