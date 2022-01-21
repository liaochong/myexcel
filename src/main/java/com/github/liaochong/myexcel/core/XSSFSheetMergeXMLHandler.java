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

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Iterator;
import java.util.Map;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * 单元格合并处理器
 *
 * @author liaochong
 * @version 1.0
 */
class XSSFSheetMergeXMLHandler extends DefaultHandler {

    private final Map<CellAddress, CellAddress> mergeCellMapping;

    public XSSFSheetMergeXMLHandler(Map<CellAddress, CellAddress> mergeCellMapping) {
        this.mergeCellMapping = mergeCellMapping;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
            return;
        }
        if ("mergeCell".equals(localName) || "x:mergeCell".equals(localName)) {
            String range = attributes.getValue("ref");
            Iterator<CellAddress> iterator = CellRangeAddress.valueOf(range).iterator();
            CellAddress firstCellAddress = null;
            while (iterator.hasNext()) {
                CellAddress cellAddress = iterator.next();
                if (firstCellAddress == null) {
                    firstCellAddress = cellAddress;
                }
                mergeCellMapping.put(cellAddress, firstCellAddress);
            }
        }
    }
}
