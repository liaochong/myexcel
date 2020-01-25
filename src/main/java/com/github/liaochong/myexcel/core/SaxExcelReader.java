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
import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.exception.ExcelReadException;
import com.github.liaochong.myexcel.exception.SaxReadException;
import com.github.liaochong.myexcel.exception.StopReadException;
import com.github.liaochong.myexcel.utils.TempFileOperator;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
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
@Slf4j
public class SaxExcelReader<T> {

    private static final int DEFAULT_SHEET_INDEX = 0;

    private List<T> result = new LinkedList<>();

    private ReadConfig<T> readConfig = new ReadConfig<>(DEFAULT_SHEET_INDEX);

    private SaxExcelReader(Class<T> dataType) {
        this.readConfig.dataType = dataType;
    }

    public static <T> SaxExcelReader<T> of(@NonNull Class<T> clazz) {
        return new SaxExcelReader<>(clazz);
    }

    public SaxExcelReader<T> sheet(Integer sheetIndex) {
        return sheets(sheetIndex);
    }

    public SaxExcelReader<T> sheet(String sheetName) {
        return sheets(sheetName);
    }

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

    public SaxExcelReader<T> charset(String charset) {
        this.readConfig.charset = charset;
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

    public List<T> read(@NonNull InputStream fileInputStream) {
        doRead(fileInputStream);
        return result;
    }

    public List<T> read(@NonNull File file) {
        doRead(file);
        return result;
    }

    public void readThen(@NonNull InputStream fileInputStream, Consumer<T> consumer) {
        this.readConfig.consumer = consumer;
        doRead(fileInputStream);
    }

    public void readThen(@NonNull File file, Consumer<T> consumer) {
        this.readConfig.consumer = consumer;
        doRead(file);
    }

    public void readThen(@NonNull InputStream fileInputStream, Function<T, Boolean> function) {
        this.readConfig.function = function;
        doRead(fileInputStream);
    }

    public void readThen(@NonNull File file, Function<T, Boolean> function) {
        this.readConfig.function = function;
        doRead(file);
    }

    private void doRead(InputStream fileInputStream) {
        Path path = this.convertToFile(fileInputStream);
        try {
            doRead(path.toFile());
        } finally {
            TempFileOperator.deleteTempFile(path);
        }
    }

    private Path convertToFile(InputStream is) {
        Path tempFile = TempFileOperator.createTempFile("i_t", Constants.XLSX);
        try (FileOutputStream fos = new FileOutputStream(tempFile.toFile())) {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = is.read(buffer)) > 0) {
                fos.write(buffer, 0, len);
            }
        } catch (IOException e) {
            TempFileOperator.deleteTempFile(tempFile);
            throw new SaxReadException("Fail to convert file inputStream to temp file", e);
        }
        return tempFile;
    }

    private void doRead(File file) {
        FileMagic fm;
        try (InputStream is = FileMagic.prepareToCheckMagic(new FileInputStream(file))) {
            fm = FileMagic.valueOf(is);
        } catch (Throwable throwable) {
            throw new SaxReadException("Fail to get excel magic", throwable);
        }
        try {
            switch (fm) {
                case OOXML:
                    doReadXlsx(file);
                    break;
                case OLE2:
                    doReadXls(file);
                    break;
                default:
                    doReadCsv(file);
            }
        } catch (Throwable e) {
            throw new SaxReadException("Fail to read excel", e);
        }
    }

    private void doReadXls(File file) {
        try {
            new HSSFSaxReadHandler<>(file, result, readConfig).process();
        } catch (StopReadException e) {
            // do nothing
        } catch (IOException e) {
            throw new SaxReadException("Fail to read xls file:" + file.getName(), e);
        }
    }

    private void doReadXlsx(File file) {
        try (OPCPackage p = OPCPackage.open(file, PackageAccess.READ)) {
            process(p);
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
        StringsCache stringsCache = new StringsCache();
        try {
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(xlsxPackage, stringsCache);
            XSSFReader xssfReader = new XSSFReader(xlsxPackage);
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            if (!readConfig.sheetNames.isEmpty()) {
                while (iter.hasNext()) {
                    try (InputStream stream = iter.next()) {
                        if (readConfig.sheetNames.contains(iter.getSheetName())) {
                            processSheet(strings, new XSSFSaxReadHandler<>(result, readConfig), stream);
                        }
                    }
                }
            } else {
                int index = 0;
                while (iter.hasNext()) {
                    try (InputStream stream = iter.next()) {
                        if (readConfig.sheetIndexs.contains(index)) {
                            processSheet(strings, new XSSFSaxReadHandler<>(result, readConfig), stream);
                        }
                        ++index;
                    }
                }
            }
        } finally {
            stringsCache.clearAll();
        }
        log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param strings          The table of strings that may be referenced by cells in the sheet
     * @param sheetInputStream The stream to read the sheet-data from.
     * @throws java.io.IOException An IO exception from the parser,
     *                             possibly from a byte stream or character stream
     *                             supplied by the application.
     * @throws SAXException        if parsing the XML data fails.
     */
    private void processSheet(
            SharedStrings strings,
            XSSFSheetXMLHandler.SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    null, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    @Getter
    @Setter
    @FieldDefaults(level = AccessLevel.PRIVATE)
    public static class ReadConfig<T> {

        Class<T> dataType;

        Set<String> sheetNames = new HashSet<>();

        Set<Integer> sheetIndexs = new HashSet<>();

        Consumer<T> consumer;

        Function<T, Boolean> function;

        Predicate<Row> rowFilter = row -> true;

        Predicate<T> beanFilter = bean -> true;

        BiFunction<Throwable, ReadContext, Boolean> exceptionFunction = (t, c) -> false;

        String charset = "UTF-8";

        Function<String, String> trim = v -> {
            if (v == null) {
                return v;
            }
            return v.trim();
        };

        public ReadConfig(int sheetIndex) {
            sheetIndexs.add(sheetIndex);
        }
    }
}
