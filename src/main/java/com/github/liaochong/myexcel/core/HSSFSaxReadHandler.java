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
import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
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
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellAddress;
import org.slf4j.Logger;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * HSSF sax处理
 *
 * @author liaochong
 * @version 1.0
 */
class HSSFSaxReadHandler<T> extends AbstractReadHandler<T> implements HSSFListener {

    private static final Logger log = org.slf4j.LoggerFactory.getLogger(HSSFSaxReadHandler.class);

    private String sheetName;

    private final POIFSFileSystem fs;

    private int lastRowNumber = -1;

    /**
     * Should we output the formula, or the value it has?
     */
    private final boolean outputFormulaValues = true;

    /**
     * For parsing Formulas
     */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

    // Records we pick up as we process
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    /**
     * So we known which sheet we're on
     */
    private int sheetIndex = -1;
    private BoundSheetRecord[] orderedBSRs;
    private final List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    // For handling formulas with string results
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;

    private final Map<Integer, Map<CellAddress, CellAddress>> mergeCellIndexMapping;

    private Map<CellAddress, String> mergeFirstCellMapping;

    private boolean detectedMergeOfThisSheet;

    private long waitCount = 0;

    private final HSSFPreReadHandler.HSSFPreData hssfPreData;

    public HSSFSaxReadHandler(File file,
                              List<T> result,
                              SaxExcelReader.ReadConfig<T> readConfig,
                              HSSFPreReadHandler.HSSFPreData hssfPreData) throws IOException {
        super(false, result, readConfig, Collections.emptyMap());
        this.fs = new POIFSFileSystem(Files.newInputStream(file.toPath()));
        this.hssfPreData = hssfPreData;
        this.mergeCellIndexMapping = hssfPreData != null ? hssfPreData.mergeCellIndexMapping : Collections.emptyMap();
    }

    public void process() throws IOException {
        long startTime = System.currentTimeMillis();
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);

        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }

        factory.processWorkbookEvents(request, fs);
        log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
    }

    @Override
    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;

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
                    if (hssfPreData != null) {
                        hssfPreData.hyperlinkMapping.remove(sheetIndex);
                    }
                    sheetIndex++;
                    setRecordAsNull();
                    lastRowNumber = -1;
                    titles = new LinkedHashMap<>();
                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                    sheetName = orderedBSRs[sheetIndex].getSheetname();
                    readContext.readConfig.startSheetConsumer.accept(sheetName, sheetIndex);
                    mergeCellMapping = mergeCellIndexMapping.getOrDefault(sheetIndex, Collections.emptyMap());
                    detectedMergeOfThisSheet = !mergeCellMapping.isEmpty();
                    waitCount = 0;
                    mergeFirstCellMapping = mergeCellMapping.values().stream().distinct().collect(Collectors.toMap(cellAddress -> cellAddress, c -> ""));
                    setFieldHandlerFunction();
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;

            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;

                thisRow = brec.getRow();
                thisColumn = brec.getColumn();
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;

                thisRow = berec.getRow();
                thisColumn = berec.getColumn();
                thisStr = berec.isBoolean() ? String.valueOf(berec.getBooleanValue()) : null;
                break;

            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;
                thisRow = frec.getRow();
                thisColumn = frec.getColumn();

                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                        nextRow = frec.getRow();
                        nextColumn = frec.getColumn();
                    } else {
                        thisStr = formatListener.formatNumberDateCell(frec);
                    }
                } else {
                    thisStr = HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression());
                }
                break;
            case StringRecord.sid:
                if (outputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord) record;
                    thisStr = srec.getString();
                    thisRow = nextRow;
                    thisColumn = nextColumn;
                    outputNextStringRecord = false;
                }
                break;

            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;

                thisRow = lrec.getRow();
                thisColumn = lrec.getColumn();
                thisStr = lrec.getValue();
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord) record;

                thisRow = lsrec.getRow();
                thisColumn = lsrec.getColumn();
                if (sstRecord != null) {
                    thisStr = sstRecord.getString(lsrec.getSSTIndex()).toString();
                }
                break;
            case NoteRecord.sid:
                NoteRecord nrec = (NoteRecord) record;

                thisRow = nrec.getRow();
                thisColumn = nrec.getColumn();
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;

                thisRow = numrec.getRow();
                thisColumn = numrec.getColumn();

                // Format
                thisStr = formatListener.formatNumberDateCell(numrec);
                break;
            case RKRecord.sid:
                RKRecord rkrec = (RKRecord) record;

                thisRow = rkrec.getRow();
                thisColumn = rkrec.getColumn();
                break;
            default:
                break;
        }

        // Handle missing column
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            thisRow = mc.getRow();
            thisColumn = mc.getColumn();
            thisStr = null;
        }

        if (record instanceof LastCellOfRowDummyRecord) {
            LastCellOfRowDummyRecord lc = (LastCellOfRowDummyRecord) record;
            thisRow = lc.getRow();
        }

        // Handle new row
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastRowNumber = thisRow;
            newRow(thisRow, !detectedMergeOfThisSheet || waitCount == 0);
            if (detectedMergeOfThisSheet && waitCount == 0) {
                waitCount = mergeCellMapping.entrySet().stream().filter(c -> c.getValue().getColumn() == 0
                        && Objects.equals(c.getValue().getRow(), lastRowNumber)
                        && c.getKey().getRow() != c.getValue().getRow()).count() + 1;
            }
            waitCount--;
        }
        boolean isSelectedSheet = this.isSelectedSheet();
        if (isSelectedSheet && thisColumn != -1) {
            if (readContext.readConfig.detectedMerge) {
                CellAddress cellAddress = new CellAddress(thisRow, thisColumn);
                String finalThisStr = thisStr;
                mergeFirstCellMapping.computeIfPresent(cellAddress, (k, v) -> finalThisStr);
                CellAddress firstCellAddress = mergeCellIndexMapping.getOrDefault(sheetIndex, Collections.emptyMap()).get(cellAddress);
                if (firstCellAddress != null) {
                    thisStr = mergeFirstCellMapping.get(firstCellAddress);
                }
            }
            if (hssfPreData != null && !hssfPreData.hyperlinkMapping.isEmpty()) {
                Map<CellAddress, Hyperlink> hyperlinkMapping = hssfPreData.hyperlinkMapping.get(sheetIndex);
                if (hyperlinkMapping != null) {
                    CellAddress cellAddress = new CellAddress(thisRow, thisColumn);
                    Hyperlink hyperlink = hyperlinkMapping.get(cellAddress);
                    readContext.setHyperlink(hyperlink);
                }
            }
            handleField(thisColumn, thisStr);
        }
        // Handle end of row
        if (record instanceof LastCellOfRowDummyRecord) {
            if (!isSelectedSheet) {
                this.titles.clear();
                return;
            }
            if (!detectedMergeOfThisSheet || waitCount == 0) {
                handleResult();
            }
        }
    }

    private boolean isSelectedSheet() {
        if (readContext.readConfig.readAllSheet) {
            return true;
        }
        if (!readContext.readConfig.sheetNames.isEmpty()) {
            return readContext.readConfig.sheetNames.contains(sheetName);
        }
        return readContext.readConfig.sheetIndexs.contains(sheetIndex);
    }
}
