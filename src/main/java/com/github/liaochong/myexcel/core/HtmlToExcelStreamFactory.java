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

import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.core.parser.Tr;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;

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

    /**
     * 线程池
     */
    private ExecutorService executorService;

    public HtmlToExcelStreamFactory(int waitSize, ExecutorService executorService) {
        this.trWaitQueue = new ArrayBlockingQueue<>(waitSize);
        this.executorService = executorService;
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
        if (Objects.isNull(executorService)) {
            Thread thread = new Thread(this::receive);
            thread.setName("Excel-builder-1");
            thread.start();
        } else {
            CompletableFuture.runAsync(this::receive, executorService);
        }
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
            e.printStackTrace();
        }
    }

    private void receive() {
        List<Tr> trList = this.getTrListFromQueue();
        int appendSize = 0;
        try {
            while (trList != STOP_FLAG_LIST) {
                log.info("Received data size:{},current waiting queue size:{}", trList.size(), trWaitQueue.size());
                for (Tr tr : trList) {
                    if (rowNum == maxRowCountOfSheet) {
                        sheetNum++;
                        this.setColWidth(colWidthMap, sheet);
                        colWidthMap = null;
                        sheet = workbook.createSheet(sheetName + " " + sheetNum);
                        rowNum = 0;
                    }
                    tr.setIndex(rowNum);
                    tr.getTdList().forEach(td -> {
                        td.setRow(rowNum);
                        td.setRowBound(rowNum);
                    });
                    rowNum++;
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
        this.setColWidth(colWidthMap, sheet);
        this.freezePane(0, sheet);
        log.info("Build Excel success,takes {} ms", System.currentTimeMillis() - startTime);
        return workbook;
    }
}
