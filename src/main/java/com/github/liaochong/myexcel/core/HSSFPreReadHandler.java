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

import com.github.liaochong.myexcel.core.context.Hyperlink;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.HyperlinkRecord;
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
class HSSFPreReadHandler extends AbstractHSSFReadHandler {

    private final HSSFPreData hssfPreData = new HSSFPreData();

    public HSSFPreReadHandler(File file,
                              SaxExcelReader.ReadConfig<?> readConfig) throws IOException {
        this.readConfig = readConfig;
        this.fs = new POIFSFileSystem(new FileInputStream(file));
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
                Map<CellAddress, CellAddress> mergeCellMapping = hssfPreData.mergeCellIndexMapping.computeIfAbsent(sheetIndex, k -> new HashMap<>());
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
                break;
            case HyperlinkRecord.sid:
                HyperlinkRecord hr = (HyperlinkRecord) record;
                Map<CellAddress, Hyperlink> hyperlinkMapping = hssfPreData.hyperlinkMapping.computeIfAbsent(sheetIndex, k -> new HashMap<>());
                hyperlinkMapping.put(new CellAddress(hr.getFirstRow(), hr.getLastColumn()), new Hyperlink(hr.getAddress(), hr.getLabel(), hr));
                hssfPreData.hyperlinkMapping.put(sheetIndex, hyperlinkMapping);
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

    public HSSFPreData getHssfPreData() {
        return hssfPreData;
    }

    public static class HSSFPreData {
        private Map<Integer, Map<CellAddress, CellAddress>> mergeCellIndexMapping = new HashMap<>();

        private Map<Integer, Map<CellAddress, Hyperlink>> hyperlinkMapping = new HashMap<>();

        public Map<Integer, Map<CellAddress, CellAddress>> getMergeCellIndexMapping() {
            return mergeCellIndexMapping;
        }

        public void setMergeCellIndexMapping(Map<Integer, Map<CellAddress, CellAddress>> mergeCellIndexMapping) {
            this.mergeCellIndexMapping = mergeCellIndexMapping;
        }

        public Map<Integer, Map<CellAddress, Hyperlink>> getHyperlinkMapping() {
            return hyperlinkMapping;
        }

        public void setHyperlinkMapping(Map<Integer, Map<CellAddress, Hyperlink>> hyperlinkMapping) {
            this.hyperlinkMapping = hyperlinkMapping;
        }
    }
}
