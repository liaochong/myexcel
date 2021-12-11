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
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
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
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
public class HSSFMetaDataSaxReadHandler implements HSSFListener {

    private BoundSheetRecord[] orderedBSRs;

    private final List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    private final POIFSFileSystem fs;

    private int lastRowNumber = -1;

    /**
     * For parsing Formulas
     */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

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
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        FormatTrackingHSSFListener formatListener = new FormatTrackingHSSFListener(listener);

        HSSFRequest request = new HSSFRequest();
        workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
        request.addListenerForAllRecords(workbookBuildingListener);
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
                    if (workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                    }
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
                BlankRecord brec = (BlankRecord) record;
                thisRow = brec.getRow();
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;
                thisRow = berec.getRow();
                break;
            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;
                thisRow = frec.getRow();
                break;
            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;
                thisRow = lrec.getRow();
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord) record;
                thisRow = lsrec.getRow();
                break;
            case NoteRecord.sid:
                NoteRecord nrec = (NoteRecord) record;
                thisRow = nrec.getRow();
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;
                thisRow = numrec.getRow();
                break;
            case RKRecord.sid:
                RKRecord rkrec = (RKRecord) record;
                thisRow = rkrec.getRow();
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
