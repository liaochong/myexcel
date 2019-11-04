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

import com.github.liaochong.myexcel.core.converter.ReadConverterContext;
import com.github.liaochong.myexcel.exception.StopReadException;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import lombok.extern.slf4j.Slf4j;
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

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * HSSF sax处理
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
class HSSFSaxHandler<T> implements HSSFListener {

    private final Map<Integer, Field> fieldMap;

    private List<T> result;

    private T obj;

    private Class<T> dataType;

    private Consumer<T> consumer;

    private Function<T, Boolean> function;

    private Predicate<Row> rowFilter;

    private Predicate<T> beanFilter;

    private Row currentRow;

    private Set<Integer> sheetIndexs;

    private String sheetName;

    private SaxExcelReader.ReadConfig<T> readConfig;

    private POIFSFileSystem fs;

    private int lastRowNumber = -1;

    /**
     * Should we output the formula, or the value it has?
     */
    private boolean outputFormulaValues = true;

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
    private List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    // For handling formulas with string results
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;

    private BiFunction<Throwable, ReadContext, Boolean> exceptionFunction;

    public HSSFSaxHandler(File file,
                          List<T> result,
                          SaxExcelReader.ReadConfig<T> readConfig) throws IOException {
        this.fs = new POIFSFileSystem(new FileInputStream(file));
        this.sheetIndexs = readConfig.getSheetIndexs();
        this.dataType = readConfig.getDataType();
        this.fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        this.result = result;
        this.consumer = readConfig.getConsumer();
        this.function = readConfig.getFunction();
        this.rowFilter = readConfig.getRowFilter();
        this.beanFilter = readConfig.getBeanFilter();
        this.readConfig = readConfig;
        this.exceptionFunction = readConfig.getExceptionFunction();
    }

    public HSSFSaxHandler(InputStream inputStream,
                          List<T> result,
                          SaxExcelReader.ReadConfig<T> readConfig) throws IOException {
        this.fs = new POIFSFileSystem(inputStream);
        this.sheetIndexs = readConfig.getSheetIndexs();
        this.dataType = readConfig.getDataType();
        this.fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        this.result = result;
        this.consumer = readConfig.getConsumer();
        this.function = readConfig.getFunction();
        this.rowFilter = readConfig.getRowFilter();
        this.beanFilter = readConfig.getBeanFilter();
        this.readConfig = readConfig;
        this.exceptionFunction = readConfig.getExceptionFunction();
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
                    sheetIndex++;
                    obj = null;
                    lastRowNumber = -1;
                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                    sheetName = orderedBSRs[sheetIndex].getSheetname();
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;

            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;

                thisRow = brec.getRow();
                thisColumn = brec.getColumn();
                thisStr = null;
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
                if (sstRecord == null) {
                    thisStr = null;
                } else {
                    thisStr = sstRecord.getString(lsrec.getSSTIndex()).toString();
                }
                break;
            case NoteRecord.sid:
                NoteRecord nrec = (NoteRecord) record;

                thisRow = nrec.getRow();
                thisColumn = nrec.getColumn();
                thisStr = null;
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
                thisStr = null;
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

        // Handle new row
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastRowNumber = thisRow;
            currentRow = new Row(thisRow);
            if (!rowFilter.test(currentRow)) {
                return;
            }
            try {
                obj = dataType.newInstance();
            } catch (InstantiationException | IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }

        if (obj == null) {
            return;
        }

        if (thisStr != null) {
            Field field = fieldMap.get(thisColumn);
            if (field == null) {
                return;
            }
            ReadContext<T> context = new ReadContext<>(obj, field, thisStr, currentRow.getRowNum(), thisColumn);
            ReadConverterContext.convert(obj, context, exceptionFunction);
        }

        // Handle end of row
        if (record instanceof LastCellOfRowDummyRecord) {
            if (!readConfig.getSheetNames().isEmpty()) {
                if (!readConfig.getSheetNames().contains(sheetName)) {
                    return;
                }
            } else if (!sheetIndexs.contains(sheetIndex)) {
                return;
            }
            if (!rowFilter.test(currentRow)) {
                return;
            }
            if (!beanFilter.test(obj)) {
                return;
            }
            if (consumer != null) {
                consumer.accept(obj);
            } else if (function != null) {
                Boolean noStop = function.apply(obj);
                if (!noStop) {
                    throw new StopReadException();
                }
            } else {
                result.add(obj);
            }
        }
    }
}
