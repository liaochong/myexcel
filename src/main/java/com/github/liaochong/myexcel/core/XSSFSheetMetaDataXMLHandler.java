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

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * @author liaochong
 * @version 1.0
 */
public class XSSFSheetMetaDataXMLHandler extends DefaultHandler {
    /**
     * some sheets do not have rowNums, Excel can read them so we should try to handle them correctly as well
     */
    private int nextRowNum;

    private int rowNum = -1;

    private final SheetMetaData sheetMetaData;

    public XSSFSheetMetaDataXMLHandler(SheetMetaData sheetMetaData) {
        this.sheetMetaData = sheetMetaData;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        if ("row".equals(localName)) {
            String rowNumStr = attributes.getValue("r");
            if (rowNumStr != null) {
                rowNum = Integer.parseInt(rowNumStr) - 1;
            } else {
                rowNum = nextRowNum;
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName)
            throws SAXException {
        if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
            return;
        }
        if ("row".equals(localName)) {
            // some sheets do not have rowNum set in the XML, Excel can read them so we should try to read them as well
            nextRowNum = rowNum + 1;
        }
    }

    @Override
    public void endDocument() throws SAXException {
        if (rowNum > -1) {
            sheetMetaData.setLastRowNum(rowNum + 1);
        }
    }
}
