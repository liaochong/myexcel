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
package com.github.liaochong.html2excel.core;

import com.github.liaochong.html2excel.core.parser.Table;
import com.github.liaochong.html2excel.core.parser.Tr;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;

/**
 * HtmlToExcelStreamFactory 流工厂
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
class HtmlToExcelStreamFactory extends AbstractExcelFactory {

    private static final int MAX_ROW_COUNT = 1048576;

    static final int DEFAULT_WAIT_SIZE = Runtime.getRuntime().availableProcessors();

    private static final List<Tr> STOP_FLAG_LIST = new ArrayList<>();

    private Sheet sheet;

    private BlockingQueue<List<Tr>> trWaitQueue;

    private boolean stop;

    private boolean exception;

    private long startTime;

    private String sheetName = "Sheet";

    private Map<Integer, Integer> colWidthMap;

    private int rowNum;

    private int sheetNum;

    public HtmlToExcelStreamFactory(int waitSize) {
        this.trWaitQueue = new ArrayBlockingQueue<>(waitSize);
    }

    public void start(Table table) {
        log.info("Start streaming building excel");
        startTime = System.currentTimeMillis();

        if (Objects.isNull(this.workbook)) {
            workbookType(WorkbookType.SXLSX);
        }
        initDefaultCellStyleMap();
        if (Objects.nonNull(table)) {
            sheetName = Objects.isNull(table.getCaption()) || table.getCaption().length() < 1 ? sheetName : table.getCaption();
        }
        this.sheet = this.workbook.createSheet(sheetName);

        CompletableFuture.runAsync(this::receive);
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
                    if (rowNum == MAX_ROW_COUNT) {
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
                colWidthMap = this.getColMaxWidthMap(trList);
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
