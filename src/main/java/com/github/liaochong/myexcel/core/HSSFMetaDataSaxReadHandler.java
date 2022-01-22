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
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NoteRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.RKRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * xls文件元信息解析器
 *
 * @author liaochong
 * @version 1.0
 */
class HSSFMetaDataSaxReadHandler implements HSSFListener {

    private BoundSheetRecord[] orderedBSRs;

    private final List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    private final POIFSFileSystem fs;

    private int lastRowNumber = -1;
    /**
     * So we known which sheet we're on
     */
    private int sheetIndex = -1;

    private final WorkbookMetaData workbookMetaData;

    public HSSFMetaDataSaxReadHandler(File file, WorkbookMetaData workbookMetaData) throws IOException {
        this.fs = new POIFSFileSystem(new FileInputStream(file));
        this.workbookMetaData = workbookMetaData;
    }

    public void process() throws IOException {
        HSSFRequest request = new HSSFRequest();
        request.addListenerForAllRecords(new EventWorkbookBuilder.SheetRecordCollectingListener(this));
        new HSSFEventFactory().processWorkbookEvents(request, fs);
        // 处理最后一个sheet
        if (lastRowNumber > -1) {
            workbookMetaData.getSheetMetaDataList().get(sheetIndex).setLastRowNum(lastRowNumber + 1);
        }
    }

    @Override
    public void processRecord(Record record) {
        int thisRow = -1;
        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add((BoundSheetRecord) record);
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    List<SheetMetaData> sheetMetaDataList = workbookMetaData.getSheetMetaDataList();
                    if (lastRowNumber > -1) {
                        sheetMetaDataList.get(sheetIndex).setLastRowNum(lastRowNumber + 1);
                    }
                    sheetIndex++;
                    lastRowNumber = -1;
                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                    String sheetName = orderedBSRs[sheetIndex].getSheetname();
                    SheetMetaData sheetMetaData = new SheetMetaData(sheetName, sheetIndex);
                    sheetMetaDataList.add(sheetMetaData);
                    workbookMetaData.setSheetCount(sheetIndex + 1);
                }
                break;
            case BlankRecord.sid:
                thisRow = ((BlankRecord) record).getRow();
                break;
            case BoolErrRecord.sid:
                thisRow = ((BoolErrRecord) record).getRow();
                break;
            case FormulaRecord.sid:
                thisRow = ((FormulaRecord) record).getRow();
                break;
            case LabelRecord.sid:
                thisRow = ((LabelRecord) record).getRow();
                break;
            case LabelSSTRecord.sid:
                thisRow = ((LabelSSTRecord) record).getRow();
                break;
            case NoteRecord.sid:
                thisRow = ((NoteRecord) record).getRow();
                break;
            case NumberRecord.sid:
                thisRow = ((NumberRecord) record).getRow();
                break;
            case RKRecord.sid:
                thisRow = ((RKRecord) record).getRow();
                break;
            default:
                break;
        }
        if (record instanceof LastCellOfRowDummyRecord) {
            LastCellOfRowDummyRecord lc = (LastCellOfRowDummyRecord) record;
            thisRow = lc.getRow();
        }
        // Handle new row
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastRowNumber = thisRow;
        }
    }
}
