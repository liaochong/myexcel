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

import com.github.liaochong.myexcel.core.cache.StringsCache;
import com.github.liaochong.myexcel.core.context.Hyperlink;
import com.github.liaochong.myexcel.core.context.ReadContext;
import com.github.liaochong.myexcel.core.context.RowContext;
import com.github.liaochong.myexcel.exception.ExcelReadException;
import com.github.liaochong.myexcel.exception.SaxReadException;
import com.github.liaochong.myexcel.exception.StopReadException;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.hibernate.validator.HibernateValidator;
import org.slf4j.Logger;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * sax模式读取excel，支持xls、xlsx、csv格式读取
 *
 * @author liaochong
 * @version 1.0
 */
public class SaxExcelReader<T> {

    private static final int DEFAULT_SHEET_INDEX = 0;
    private static final Logger log = org.slf4j.LoggerFactory.getLogger(SaxExcelReader.class);
    /**
     * 元数据
     */
    private WorkbookMetaData workbookMetaData;

    private final List<T> result = new LinkedList<>();

    private final ReadConfig<T> readConfig = new ReadConfig<>(DEFAULT_SHEET_INDEX);

    private SaxExcelReader(Class<T> dataType) {
        this.readConfig.dataType = dataType;
    }

    public static <T> SaxExcelReader<T> of(Class<T> clazz) {
        return new SaxExcelReader<>(clazz);
    }

    public SaxExcelReader<T> sheet(Integer sheetIndex) {
        return sheets(sheetIndex);
    }

    public SaxExcelReader<T> sheet(String sheetName) {
        return sheets(sheetName);
    }

    public Validator validator;

    public SaxExcelReader<T> sheets(Integer... sheetIndexs) {
        this.readConfig.sheetIndexs.clear();
        this.readConfig.sheetIndexs.addAll(Arrays.asList(sheetIndexs));
        return this;
    }

    public SaxExcelReader<T> sheets(String... sheetNames) {
        this.readConfig.sheetNames.clear();
        this.readConfig.sheetNames.addAll(Arrays.asList(sheetNames));
        return this;
    }

    public SaxExcelReader<T> rowFilter(Predicate<Row> rowFilter) {
        this.readConfig.rowFilter = rowFilter;
        return this;
    }

    public SaxExcelReader<T> beanFilter(Predicate<T> beanFilter) {
        this.readConfig.beanFilter = beanFilter;
        return this;
    }

    @Deprecated
    public SaxExcelReader<T> charset(String charset) {
        this.readConfig.csvCharset = charset;
        return this;
    }

    public SaxExcelReader<T> csvCharset(String charset) {
        this.readConfig.csvCharset = charset;
        return this;
    }

    public SaxExcelReader<T> csvDelimiter(char delimiter) {
        this.readConfig.csvDelimiter = delimiter;
        return this;
    }

    public SaxExcelReader<T> exceptionally(BiFunction<Throwable, ReadContext, Boolean> exceptionFunction) {
        this.readConfig.exceptionFunction = exceptionFunction;
        return this;
    }

    public SaxExcelReader<T> noTrim() {
        this.readConfig.trim = v -> v;
        return this;
    }

    public SaxExcelReader<T> trimToNull() {
        this.readConfig.trim = StringUtil::trimToNull;
        return this;
    }

    public SaxExcelReader<T> readAllSheet() {
        this.readConfig.readAllSheet = true;
        return this;
    }

    public SaxExcelReader<T> ignoreBlankRow() {
        this.readConfig.ignoreBlankRow = true;
        return this;
    }

    public SaxExcelReader<T> stopReadingOnBlankRow() {
        this.readConfig.stopReadingOnBlankRow = true;
        return this;
    }

    public SaxExcelReader<T> startSheet(BiConsumer<String, Integer> startSheetConsumer) {
        this.readConfig.startSheetConsumer = startSheetConsumer;
        return this;
    }

    public SaxExcelReader<T> detectedMerge() {
        this.readConfig.detectedMerge = true;
        return this;
    }

    public List<T> read(Path path) {
        doRead(path.toFile());
        return result;
    }

    public List<T> read(InputStream fileInputStream) {
        doRead(fileInputStream);
        return result;
    }

    public ValidationObject<T> validRead(InputStream fileInputStream) {
        ValidationObject<T> validationObject = new ValidationObject<>();
        this.readThen(fileInputStream, (t, rowContext) -> {
            doValidRead(t, rowContext, validationObject);
        });
        return validationObject;
    }

    public List<T> read(File file) {
        doRead(file);
        return result;
    }

    public ValidationObject<T> validRead(File file) {
        ValidationObject<T> validationObject = new ValidationObject<>();
        this.readThen(file, (t, rowContext) -> {
            doValidRead(t, rowContext, validationObject);
        });
        return validationObject;
    }

    private void doValidRead(T t, RowContext rowContext, ValidationObject<T> validationObject) {
        validationObject.getObjects().add(t);
        Set<ConstraintViolation<T>> violations = getValidator().validate(t, t.getClass());
        ValidationObject.ValidationInfo<T> validationInfo = new ValidationObject.ValidationInfo<>();
        validationInfo.setRowNum(rowContext.getRowNum());
        validationInfo.setConstraintViolations(violations);
        validationObject.getValidationInfos().add(validationInfo);
    }

    public void readThen(InputStream fileInputStream, Consumer<T> consumer) {
        this.readConfig.consumer = consumer;
        doRead(fileInputStream);
    }

    public void readThen(File file, Consumer<T> consumer) {
        this.readConfig.consumer = consumer;
        doRead(file);
    }

    public void readThen(Path path, Consumer<T> consumer) {
        this.readConfig.consumer = consumer;
        doRead(path.toFile());
    }

    public void readThen(InputStream fileInputStream, BiConsumer<T, RowContext> contextConsumer) {
        this.readConfig.contextConsumer = contextConsumer;
        doRead(fileInputStream);
    }

    public void readThen(File file, BiConsumer<T, RowContext> contextConsumer) {
        this.readConfig.contextConsumer = contextConsumer;
        doRead(file);
    }

    public void readThen(Path path, BiConsumer<T, RowContext> contextConsumer) {
        this.readConfig.contextConsumer = contextConsumer;
        doRead(path.toFile());
    }

    public void readThen(File file, BiFunction<T, RowContext, Boolean> contextFunction) {
        this.readConfig.contextFunction = contextFunction;
        doRead(file);
    }

    public void readThen(InputStream fileInputStream, BiFunction<T, RowContext, Boolean> contextFunction) {
        this.readConfig.contextFunction = contextFunction;
        doRead(fileInputStream);
    }

    public void readThen(Path path, BiFunction<T, RowContext, Boolean> contextFunction) {
        this.readConfig.contextFunction = contextFunction;
        doRead(path.toFile());
    }

    public void readThen(InputStream fileInputStream, Function<T, Boolean> function) {
        this.readConfig.function = function;
        doRead(fileInputStream);
    }

    public void readThen(File file, Function<T, Boolean> function) {
        this.readConfig.function = function;
        doRead(file);
    }

    public void readThen(Path path, Function<T, Boolean> function) {
        this.readConfig.function = function;
        doRead(path.toFile());
    }

    public static WorkbookMetaData getWorkbookMetaData(Path path) {
        SaxExcelReader<Void> saxExcelReader = new SaxExcelReader<>(null);
        saxExcelReader.doRead(path.toFile(), true);
        return saxExcelReader.workbookMetaData;
    }

    public static WorkbookMetaData getWorkbookMetaData(InputStream fileInputStream) {
        SaxExcelReader<Void> saxExcelReader = new SaxExcelReader<>(null);
        saxExcelReader.doRead(fileInputStream, true);
        return saxExcelReader.workbookMetaData;
    }

    public static WorkbookMetaData getWorkbookMetaData(File file) {
        SaxExcelReader<Void> saxExcelReader = new SaxExcelReader<>(null);
        saxExcelReader.doRead(file, true);
        return saxExcelReader.workbookMetaData;
    }

    private synchronized Validator getValidator() {
        if (validator == null) {
            try (ValidatorFactory validatorFactory = Validation
                    .byProvider(HibernateValidator.class)
                    .configure()
                    .buildValidatorFactory()) {
                validator = validatorFactory.getValidator();
            }
        }
        return validator;
    }

    private void doRead(InputStream fileInputStream) {
        this.doRead(fileInputStream, false);
    }

    private void doRead(InputStream fileInputStream, boolean readMetaData) {
        Path path = TempFileOperator.convertToFile(fileInputStream);
        try {
            doRead(path.toFile(), readMetaData);
        } finally {
            TempFileOperator.deleteTempFile(path);
        }
    }

    private void doRead(File file) {
        this.doRead(file, false);
    }

    private void doRead(File file, boolean readMetaData) {
        FileMagic fm;
        try (InputStream is = FileMagic.prepareToCheckMagic(new FileInputStream(file))) {
            fm = FileMagic.valueOf(is);
        } catch (Throwable throwable) {
            throw new SaxReadException("Fail to get excel magic", throwable);
        }
        try {
            switch (fm) {
                case OOXML:
                    doReadXlsx(file, readMetaData);
                    break;
                case OLE2:
                    doReadXls(file, readMetaData);
                    break;
                default:
                    if (readMetaData) {
                        throw new UnsupportedOperationException();
                    }
                    doReadCsv(file);
            }
        } catch (Throwable e) {
            throw new SaxReadException("Fail to read excel", e);
        }
    }

    private void doReadXls(File file, boolean readMetaData) {
        try {
            if (readMetaData) {
                workbookMetaData = new WorkbookMetaData();
                new HSSFMetaDataSaxReadHandler(file, workbookMetaData).process();
            } else {
                HSSFPreReadHandler.HSSFPreData hssfPreData = null;
                if (readConfig.detectedMerge || readConfig.detectedHyperlink()) {
                    HSSFPreReadHandler hssfPreReadHandler = new HSSFPreReadHandler(file, readConfig);
                    hssfPreReadHandler.process();
                    hssfPreData = hssfPreReadHandler.getHssfPreData();
                }
                if (hssfPreData == null || hssfPreData.mergeCellIndexMapping == null) {
                    readConfig.detectedMerge = false;
                }
                new HSSFSaxReadHandler<>(file, result, readConfig, hssfPreData).process();
            }
        } catch (StopReadException e) {
            // do nothing
        } catch (IOException e) {
            throw new SaxReadException("Fail to read xls file:" + file.getName(), e);
        }
    }

    private void doReadXlsx(File file, boolean readMetaData) {
        try (OPCPackage p = OPCPackage.open(file, PackageAccess.READ)) {
            if (readMetaData) {
                processMetaData(p);
            } else {
                process(p);
            }
        } catch (StopReadException e) {
            // do nothing
        } catch (Exception e) {
            throw new SaxReadException("Fail to read xlsx file:" + file.getName(), e);
        }
    }

    private void doReadCsv(File file) {
        try {
            new CsvReadHandler<>(Files.newInputStream(file.toPath()), readConfig, result).read();
        } catch (StopReadException e) {
            // do nothing
        } catch (Throwable throwable) {
            throw new ExcelReadException("Fail to read csv file:" + file.getName(), throwable);
        }
    }

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException  If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    private void process(OPCPackage xlsxPackage) throws IOException, OpenXML4JException, SAXException {
        long startTime = System.currentTimeMillis();
        Map<Integer, XSSFSheetPreXMLHandler.XSSFPreData> preDataIndexMapping = this.processPreData(xlsxPackage);
        StringsCache stringsCache = new StringsCache();
        try {
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(xlsxPackage, stringsCache);
            this.doReadSheet(xlsxPackage, xssfReadContext -> {
                readConfig.startSheetConsumer.accept(xssfReadContext.sheetName, xssfReadContext.sheetIndex);
                XSSFSheetPreXMLHandler.XSSFPreData xssfPreData = preDataIndexMapping.get(xssfReadContext.sheetIndex);
                ContentHandler handler = new XSSFSheetXMLHandler(
                        xssfPreData, strings, new XSSFSaxReadHandler<>(result, readConfig, xssfPreData));
                processSheet(handler, xssfReadContext.inputStream);
                preDataIndexMapping.remove(xssfReadContext.sheetIndex);
            });
        } finally {
            stringsCache.clearAll();
        }
        log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
    }

    private Map<Integer, XSSFSheetPreXMLHandler.XSSFPreData> processPreData(OPCPackage xlsxPackage) throws IOException, OpenXML4JException {
        if (!readConfig.detectedMerge && !readConfig.detectedHyperlink()) {
            return Collections.emptyMap();
        }
        Map<Integer, XSSFSheetPreXMLHandler.XSSFPreData> preDataIndexMapping = new HashMap<>();
        this.doReadSheet(xlsxPackage, xssfReadContext -> {
            xssfReadContext.packageRelationshipCollection = Optional.ofNullable(xssfReadContext.sheetIterator.getSheetPart())
                    .map(packagePart -> {
                        try {
                            return packagePart.getRelationships();
                        } catch (InvalidFormatException e) {
                            throw new SaxReadException("Get relationships failure.", e);
                        }
                    }).orElse(null);
            XSSFSheetPreXMLHandler xssfSheetPreXMLHandler = new XSSFSheetPreXMLHandler(readConfig, xssfReadContext);
            processSheet(xssfSheetPreXMLHandler, xssfReadContext.inputStream);
            preDataIndexMapping.put(xssfReadContext.sheetIndex, xssfSheetPreXMLHandler.getXssfPreData());
        });
        if (preDataIndexMapping.isEmpty()) {
            readConfig.detectedMerge = false;
        }
        return preDataIndexMapping;
    }

    private void processMetaData(OPCPackage xlsxPackage) throws IOException, OpenXML4JException {
        workbookMetaData = new WorkbookMetaData();
        readConfig.readAllSheet = true;
        int lastIndex = this.doReadSheet(xlsxPackage, xssfReadContext -> {
            SheetMetaData sheetMetaData = new SheetMetaData(xssfReadContext.sheetName, xssfReadContext.sheetIndex);
            this.processSheet(new XSSFSheetMetaDataXMLHandler(sheetMetaData), xssfReadContext.inputStream);
            // 设置元数据信息
            workbookMetaData.getSheetMetaDataList().add(sheetMetaData);
        });
        if (lastIndex > -1) {
            workbookMetaData.setSheetCount(lastIndex + 1);
        }
    }

    private int doReadSheet(OPCPackage xlsxPackage, Consumer<XSSFReadContext> consumer) throws IOException, OpenXML4JException {
        XSSFReader.SheetIterator iter = this.getSheetIterator(xlsxPackage);
        Function<XSSFReadContext, Boolean> acceptFunction = this.getSheetAcceptFunction();
        int index = -1;
        while (iter.hasNext()) {
            ++index;
            try (InputStream stream = iter.next()) {
                XSSFReadContext xssfReadContext = new XSSFReadContext(iter, stream, index, iter.getSheetName());
                if (acceptFunction.apply(xssfReadContext)) {
                    consumer.accept(xssfReadContext);
                }
            }
        }
        return index;
    }

    private XSSFReader.SheetIterator getSheetIterator(OPCPackage xlsxPackage) throws IOException, OpenXML4JException {
        XSSFReader xssfReader = new XSSFReader(xlsxPackage);
        return (XSSFReader.SheetIterator) xssfReader.getSheetsData();
    }

    private Function<XSSFReadContext, Boolean> getSheetAcceptFunction() {
        Function<XSSFReadContext, Boolean> acceptFunction = xssfReadContext -> true;
        if (readConfig.readAllSheet) {
            acceptFunction = xssfReadContext -> true;
        } else if (!readConfig.sheetNames.isEmpty()) {
            acceptFunction = xssfReadContext -> readConfig.sheetNames.contains(xssfReadContext.sheetName);
        } else if (!readConfig.sheetIndexs.isEmpty()) {
            acceptFunction = xssfReadContext -> readConfig.sheetIndexs.contains(xssfReadContext.sheetIndex);
        }
        return acceptFunction;
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param sheetInputStream The stream to read the sheet-data from.
     */
    private void processSheet(
            ContentHandler handler,
            InputStream sheetInputStream) {
        try {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            sheetParser.setContentHandler(handler);
            sheetParser.parse(new InputSource(sheetInputStream));
        } catch (ParserConfigurationException | IOException | SAXException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    public static final class ReadConfig<T> {

        public Class<T> dataType;

        public Set<String> sheetNames = new HashSet<>();

        public Set<Integer> sheetIndexs = new HashSet<>();

        public Consumer<T> consumer;

        public BiConsumer<T, RowContext> contextConsumer;

        public Function<T, Boolean> function;

        public BiFunction<T, RowContext, Boolean> contextFunction;

        public Predicate<Row> rowFilter = row -> true;

        public Predicate<T> beanFilter = bean -> true;

        public BiFunction<Throwable, ReadContext, Boolean> exceptionFunction = (t, c) -> false;

        public String csvCharset = "UTF-8";

        public char csvDelimiter = ',';

        public Function<String, String> trim = v -> {
            if (v == null) {
                return v;
            }
            return v.trim();
        };

        public boolean readAllSheet;
        /**
         * 是否忽略空白行，默认为否
         */
        public boolean ignoreBlankRow = false;
        /**
         * 是否在遇到空白行时停止读取
         */
        public boolean stopReadingOnBlankRow = false;

        public boolean detectedMerge;

        public BiConsumer<String, Integer> startSheetConsumer = (sheetName, sheetIndex) -> {
            log.info("Start read excel, sheet:{},index:{}", sheetName, sheetIndex);
        };

        public ReadConfig(int sheetIndex) {
            sheetIndexs.add(sheetIndex);
        }

        public boolean detectedHyperlink() {
            if (dataType == Map.class) {
                return false;
            }
            return ReflectUtil.getFieldDefinitionMapOfExcelColumn(dataType).values().stream().anyMatch(f -> f.getField().getType() == Hyperlink.class);
        }
    }

    public static class XSSFReadContext {

        public XSSFReader.SheetIterator sheetIterator;
        public InputStream inputStream;
        public Integer sheetIndex;
        public String sheetName;

        public PackageRelationshipCollection packageRelationshipCollection;

        public XSSFReadContext(XSSFReader.SheetIterator sheetIterator, InputStream inputStream, Integer sheetIndex, String sheetName) {
            this.sheetIterator = sheetIterator;
            this.inputStream = inputStream;
            this.sheetIndex = sheetIndex;
            this.sheetName = sheetName;
        }
    }
}
