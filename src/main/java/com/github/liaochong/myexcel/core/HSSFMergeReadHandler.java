/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.MergeCellsRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
class HSSFMergeReadHandler extends AbstractHSSFReadHandler {

    private final Map<Integer, Map<CellAddress, CellAddress>> mergeCellIndexMapping;

    public HSSFMergeReadHandler(File file,
                                SaxExcelReader.ReadConfig<?> readConfig,
                                Map<Integer, Map<CellAddress, CellAddress>> mergeCellIndexMapping) throws IOException {
        this.readConfig = readConfig;
        this.mergeCellIndexMapping = mergeCellIndexMapping;
        this.fs = new POIFSFileSystem(new FileInputStream(file));
    }

    public void process() throws IOException {
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        request.addListenerForAllRecords(new EventWorkbookBuilder.SheetRecordCollectingListener(this));
        factory.processWorkbookEvents(request, fs);
    }

    @Override
    public void processRecord(Record record) {
        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add((BoundSheetRecord) record);
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    sheetIndex++;
                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                    sheetName = orderedBSRs[sheetIndex].getSheetname();
                }
                break;
            case MergeCellsRecord.sid:
                if (!isSelectedSheet()) {
                    return;
                }
                MergeCellsRecord mergeCellsRecord = (MergeCellsRecord) record;
                int numAreas = mergeCellsRecord.getNumAreas();
                Map<CellAddress, CellAddress> mergeCellMapping = new HashMap<>();
                for (int i = 0; i < numAreas; i++) {
                    Iterator<CellAddress> iterator = mergeCellsRecord.getAreaAt(i).iterator();
                    CellAddress firstCellAddress = null;
                    while (iterator.hasNext()) {
                        CellAddress cellAddress = iterator.next();
                        if (firstCellAddress == null) {
                            firstCellAddress = cellAddress;
                        } else {
                            mergeCellMapping.put(cellAddress, firstCellAddress);
                        }
                    }
                }
                mergeCellIndexMapping.put(sheetIndex, mergeCellMapping);
                break;
            default:
                break;
        }
    }

    private boolean isSelectedSheet() {
        if (readConfig.readAllSheet) {
            return true;
        }
        if (!readConfig.sheetNames.isEmpty()) {
            return readConfig.sheetNames.contains(sheetName);
        }
        return readConfig.sheetIndexs.contains(sheetIndex);
    }
}
