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

import lombok.NonNull;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * sax模式读取excel
 *
 * @author liaochong
 * @version 1.0
 */
public class SaxExcelReader<T> extends AbstractExcelReader<T> {

    private OPCPackage xlsxPackage;

    private Consumer<T> consumer;

    private List<T> result = Collections.emptyList();

    private SaxExcelReader(Class<T> dataType) {
        this.dataType = dataType;
    }

    public static <T> SaxExcelReader<T> of(@NonNull Class<T> clazz) {
        return new SaxExcelReader<>(clazz);
    }

    @Override
    public ExcelReader<T> sheet(int index) {
        this.sheetIndex = index;
        return this;
    }

    @Override
    public List<T> read(@NonNull InputStream fileInputStream) {
        try (OPCPackage p = OPCPackage.open(fileInputStream)) {
            xlsxPackage = p;
            process();
            return result;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public List<T> read(@NonNull File file) {
        try (OPCPackage p = OPCPackage.open(file, PackageAccess.READ)) {
            xlsxPackage = p;
            process();
            return result;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void readThen(@NonNull InputStream fileInputStream, Consumer<T> consumer) {
        try (OPCPackage p = OPCPackage.open(fileInputStream)) {
            xlsxPackage = p;
            this.consumer = consumer;
            process();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void readThen(@NonNull File file, Consumer<T> consumer) {
        try (OPCPackage p = OPCPackage.open(file, PackageAccess.READ)) {
            xlsxPackage = p;
            this.consumer = consumer;
            process();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException  If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    private void process() throws IOException, OpenXML4JException, SAXException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();

        Map<Integer, Field> fieldMap = this.getFieldMap();
        result = new LinkedList<>();
        int index = 0;
        while (iter.hasNext()) {
            if (sheetIndex > index) {
                break;
            }
            if (sheetIndex == index) {
                try (InputStream stream = iter.next()) {
                    processSheet(styles, strings, new SaxHandler<>(dataType, fieldMap, result, consumer), stream);
                }
            }
            ++index;
        }
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles           The table of styles that may be referenced by cells in the sheet
     * @param strings          The table of strings that may be referenced by cells in the sheet
     * @param sheetInputStream The stream to read the sheet-data from.
     * @throws java.io.IOException An IO exception from the parser,
     *                             possibly from a byte stream or character stream
     *                             supplied by the application.
     * @throws SAXException        if parsing the XML data fails.
     */
    private void processSheet(
            Styles styles,
            SharedStrings strings,
            XSSFSheetXMLHandler.SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

}
