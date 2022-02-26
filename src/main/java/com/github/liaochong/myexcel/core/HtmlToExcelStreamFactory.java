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
import com.github.liaochong.myexcel.core.parser.StyleParser;
import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.core.parser.Td;
import com.github.liaochong.myexcel.core.parser.Tr;
import com.github.liaochong.myexcel.exception.ExcelBuildException;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;
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
class HtmlToExcelStreamFactory extends AbstractExcelFactory {

    private static final Tr STOP_FLAG = new Tr(-1, 0);
    private static final Logger log = org.slf4j.LoggerFactory.getLogger(HtmlToExcelStreamFactory.class);

    private int maxRowCountOfSheet = XLSX_MAX_ROW_COUNT;

    private Sheet sheet;

    private boolean stop;

    private volatile boolean exception;

    private long startTime;

    private String sheetName = "Sheet";

    private Map<Integer, Integer> colWidthMap = new HashMap<>();

    private int rowNum;

    private int sheetNum;

    private int maxColIndex;

    /**
     * 计数器
     */
    private int count;

    private List<Tr> titles;
    /**
     * 临时文件
     */
    private final List<Path> tempFilePaths = new ArrayList<>();

    private final List<CompletableFuture<Void>> futures = new LinkedList<>();

    /**
     * 接收线程
     */
    volatile Thread receiveThread;

    private final HtmlToExcelStreamFactoryContext context;

    public HtmlToExcelStreamFactory(HtmlToExcelStreamFactoryContext context) {
        this.context = context;
    }

    /**
     * 消费者是否完结
     */
    private volatile boolean consumeFinished = false;

    public void start(Table table, Workbook workbook) {
        log.info("Start build excel");
        if (workbook != null) {
            this.workbook = workbook;
        }
        startTime = System.currentTimeMillis();
        if (table != null && table.caption != null) {
            sheetName = table.caption;
        }
        Thread thread = new Thread(this::receive);
        thread.setName("myexcel-exec-" + thread.getId());
        thread.start();
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
        this.putTrToQueue(tr);
    }

    private void receive() {
        try {
            if (this.workbook == null) {
                workbookType(WorkbookType.SXLSX);
            }
            if (isHssf) {
                maxRowCountOfSheet = XLS_MAX_ROW_COUNT;
            }
            initCellStyle(this.workbook);
            receiveThread = Thread.currentThread();
            this.sheet = workbook.getSheet(sheetName);
            if (this.sheet == null) {
                this.sheet = this.createSheet(sheetName);
            } else {
                rowNum = count = this.sheet.getLastRowNum() + 1;
                if (rowNum > 0) {
                    context.trWaitQueue.clear();
                }
            }
            Tr tr = this.getTrFromQueue();
            if (maxColIndex == 0) {
                int tdSize = tr.tdList.size();
                maxColIndex = tdSize > 0 ? tdSize - 1 : 0;
            }
            int totalSize = 0;
            while (tr != STOP_FLAG) {
                if (context.capacity > 0 && count == context.capacity) {
                    // 上一份数据保存
                    this.storeToTempFile();
                    // 开启下一份数据
                    this.initNewWorkbook();
                }
                createNextSheet();
                setTdStyle(tr);
                appendRow(tr);
                totalSize++;
                tr.colWidthMap.forEach((k, v) -> {
                    Integer val = this.colWidthMap.get(k);
                    if (val == null || v > val) {
                        this.colWidthMap.put(k, v);
                    }
                });
                tr = this.getTrFromQueue();
            }
            consumeFinished = true;
            log.info("Total size:{}", totalSize);
        } catch (Exception e) {
            exception = true;
            context.trWaitQueue.clear();
            context.trWaitQueue = null;
            clear();
            log.error("An exception occurred while processing", e);
            throw new ExcelBuildException("An exception occurred while processing", e);
        }
    }

    private void createNextSheet() {
        if (rowNum >= maxRowCountOfSheet) {
            sheetNum++;
            this.setColWidth(colWidthMap, sheet, maxColIndex);
            colWidthMap = new HashMap<>();
            sheet = workbook.getSheet(sheetName + " (" + sheetNum + ")");
            if (sheet == null) {
                sheet = this.createSheet(sheetName + " (" + sheetNum + ")");
                rowNum = 0;
                this.setTitles();
            } else {
                rowNum = this.sheet.getLastRowNum() + 1;
                count += rowNum;
                if (rowNum == 0) {
                    this.setTitles();
                }
                createNextSheet();
            }
        }
    }

    private void setTdStyle(Tr tr) {
        if (tr.fromTemplate) {
            return;
        }
        context.styleParser.toggle();
        // 是否为自定义宽度
        boolean isCustomWidth = !Objects.equals(tr.colWidthMap, Collections.emptyMap());
        for (int i = 0, size = tr.tdList.size(); i < size; i++) {
            Td td = tr.tdList.get(i);
            if (td.th) {
                td.style = context.styleParser.getTitleStyle("title&" + td.col);
            } else {
                td.style = context.styleParser.getCellStyle(i, td.tdContentType, td.format);
            }
            if (isCustomWidth) {
                String width = td.style.get("width");
                if (StringUtil.isNotBlank(width)) {
                    tr.colWidthMap.put(i, TdUtil.getValue(width));
                }
            }
        }
    }

    private Tr getTrFromQueue() throws InterruptedException {
        Tr tr = context.trWaitQueue.poll(15, TimeUnit.MINUTES);
        if (tr == null) {
            throw new IllegalStateException("Get tr failure,timeout 15 minutes.");
        }
        return tr;
    }

    @Override
    public Workbook build() {
        waiting();
        this.setColWidth(colWidthMap, sheet, maxColIndex);
        log.info("Build Excel success,takes {} ms", System.currentTimeMillis() - startTime);
        return workbook;
    }

    List<Path> buildAsPaths() {
        waiting();
        this.storeToTempFile();
        futures.forEach(CompletableFuture::join);
        log.info("Build Excel success,takes {} ms", System.currentTimeMillis() - startTime);
        return tempFilePaths.stream().filter(path -> Objects.nonNull(path) && path.toFile().exists()).collect(Collectors.toList());
    }

    protected void waiting() {
        if (exception) {
            throw new IllegalStateException("An exception occurred while processing");
        }
        this.stop = true;
        this.putTrToQueue(STOP_FLAG);
        while (!consumeFinished) {
            // wait all tr received
            if (exception) {
                throw new IllegalThreadStateException("An exception occurred while processing");
            }
        }
    }

    private void putTrToQueue(Tr tr) {
        try {
            boolean putSuccess = context.trWaitQueue.offer(tr, 1, TimeUnit.HOURS);
            if (!putSuccess) {
                throw new IllegalStateException("Put tr to queue failure,timeout 1 hour.");
            }
        } catch (InterruptedException e) {
            if (receiveThread != null) {
                receiveThread.interrupt();
            }
            throw new ExcelBuildException("Put tr to queue failure", e);
        }
    }

    private void storeToTempFile() {
        String suffix = isHssf ? Constants.XLS : Constants.XLSX;
        Path path = TempFileOperator.createTempFile("s_t_r_p", suffix);
        tempFilePaths.add(path);
        try {
            if (context.executorService != null) {
                Workbook tempWorkbook = workbook;
                Sheet tempSheet = sheet;
                Map<Integer, Integer> tempColWidthMap = colWidthMap;
                CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
                    this.setColWidth(tempColWidthMap, tempSheet, maxColIndex);
                    try {
                        this.createEmptySheetIfAbsent(tempWorkbook);
                        FileExportUtil.export(tempWorkbook, path.toFile());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                    if (context.pathConsumer != null) {
                        context.pathConsumer.accept(path);
                    }
                }, context.executorService);
                futures.add(future);
            } else {
                this.setColWidth(colWidthMap, sheet, maxColIndex);
                this.createEmptySheetIfAbsent(workbook);
                FileExportUtil.export(workbook, path.toFile());
                if (Objects.nonNull(context.pathConsumer)) {
                    context.pathConsumer.accept(path);
                }
            }
        } catch (IOException e) {
            clear();
            throw new RuntimeException(e);
        }
    }

    private void createEmptySheetIfAbsent(Workbook tempWorkbook) {
        if (tempWorkbook.getNumberOfSheets() == 0) {
            this.createSheet(sheetName);
        }
    }

    private void freezePane(Sheet sheet) {
        if (context.fixedTitles && titles != null) {
            sheet.createFreezePane(0, titles.size());
        }
        if (context.freezePane != null) {
            sheet.createFreezePane(context.freezePane.colSplit, context.freezePane.rowSplit);
        }
    }

    private void initNewWorkbook() {
        workbook = null;
        workbookType(isHssf ? WorkbookType.XLS : WorkbookType.SXLSX);
        sheetNum = 0;
        rowNum = 0;
        count = 0;
        colWidthMap = new HashMap<>();
        clearCache();
        initCellStyle(workbook);
        sheet = this.createSheet(sheetName);
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

    private Sheet createSheet(String sheetName) {
        Sheet sheet = workbook.createSheet(sheetName);
        this.freezePane(sheet);
        // 默认自适应打印页
        PrintSetup ps = sheet.getPrintSetup();
        ps.setFitHeight((short) 1);
        ps.setFitWidth((short) 1);
        context.startSheetConsumer.accept(sheet);
        return sheet;
    }

    private void appendRow(Tr tr) {
        tr.index = rowNum;
        tr.tdList.forEach(td -> {
            td.row = rowNum;
        });
        rowNum++;
        count++;
        this.createRow(tr, sheet);
    }

    Path buildAsZip(String fileName) {
        waiting();
        this.storeToTempFile();
        futures.forEach(CompletableFuture::join);
        String suffix = isHssf ? Constants.XLS : Constants.XLSX;
        Path zipFile = TempFileOperator.createTempFile(fileName, ".zip");
        try (ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(zipFile))) {
            for (int i = 1, size = tempFilePaths.size(); i <= size; i++) {
                Path path = tempFilePaths.get(i - 1);
                ZipEntry zipEntry = new ZipEntry(fileName + " (" + i + ")" + suffix);
                out.putNextEntry(zipEntry);
                out.write(Files.readAllBytes(path));
                out.closeEntry();
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            clear();
            tempFilePaths.clear();
        }
        tempFilePaths.add(zipFile);
        return zipFile;
    }

    public void cancel() {
        waiting();
        clear();
    }

    public void clear() {
        if (receiveThread != null && receiveThread.isAlive()) {
            receiveThread.interrupt();
        }
        closeWorkbook();
        TempFileOperator.deleteTempFiles(tempFilePaths);
    }

    /**
     * 上下文
     */
    static class HtmlToExcelStreamFactoryContext {

        BlockingQueue<Tr> trWaitQueue = new LinkedBlockingQueue<>(Runtime.getRuntime().availableProcessors() * 2);
        /**
         * 线程池
         */
        ExecutorService executorService;
        /**
         * 文件分割,excel容量
         */
        int capacity;

        Consumer<Path> pathConsumer;
        /**
         * 是否固定标题
         */
        boolean fixedTitles;

        StyleParser styleParser;

        /**
         * sheet前置处理函数
         */
        Consumer<Sheet> startSheetConsumer = sheet -> {
        };

        FreezePane freezePane;
    }
}
