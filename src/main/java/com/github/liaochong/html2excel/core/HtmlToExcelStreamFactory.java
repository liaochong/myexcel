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

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
class HtmlToExcelStreamFactory extends AbstractExcelFactory {

    private static final List<Tr> STOP_FLAG_LIST = new ArrayList<>();

    private Sheet sheet;

    private BlockingQueue<List<Tr>> trWaitQueue;

    private boolean stop;

    private long startTime;

    private String sheetName = "Sheet";

    private Map<Integer, Integer> colWidthMap;

    public HtmlToExcelStreamFactory(BlockingQueue<List<Tr>> trWaitQueue) {
        this.trWaitQueue = trWaitQueue;
    }

    public HtmlToExcelStreamFactory(int waitSize) {
        this.trWaitQueue = new ArrayBlockingQueue<>(waitSize);
    }

    public void start(Table table) {
        log.info("Start building excel");
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
        if (stop) {
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
        log.info("Start receiving");
        List<Tr> trList = this.getTrListFromQueue();
        while (trList != STOP_FLAG_LIST) {
            for (Tr tr : trList) {
                this.createRow(tr, sheet);
            }
            colWidthMap = this.getColMaxWidthMap(trList);
            trList = this.getTrListFromQueue();
        }
        log.info("End of reception");
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
        this.stop = true;
        while (!trWaitQueue.isEmpty()) {
            // wait all tr received
        }
        try {
            trWaitQueue.put(STOP_FLAG_LIST);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        this.setColWidth(colWidthMap, sheet);
        this.freezePane(0, sheet);
        log.info("Build Excel success, takes {} ms", System.currentTimeMillis() - startTime);
        return workbook;
    }
}
