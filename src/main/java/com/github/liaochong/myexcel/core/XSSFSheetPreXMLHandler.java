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

import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.context.Hyperlink;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Optional;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * 预数据处理器
 *
 * @author liaochong
 * @version 1.0
 */
class XSSFSheetPreXMLHandler extends DefaultHandler {

    private final XSSFPreData xssfPreData = new XSSFPreData();

    private final SaxExcelReader.XSSFReadContext xssfReadContext;

    public XSSFSheetPreXMLHandler(SaxExcelReader.XSSFReadContext xssfReadContext) {
        this.xssfReadContext = xssfReadContext;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        this.doProcessMerge(uri, localName, attributes);
        this.doProcessHyperlink(attributes);
    }

    private void doProcessMerge(String uri, String localName, Attributes attributes) {
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
                } else {
                    xssfPreData.mergeCellMapping.put(cellAddress, firstCellAddress);
                }
            }
        }
    }

    private void doProcessHyperlink(Attributes attributes) {
        String ref = attributes.getValue(Constants.ATTRIBUTE_REF);
        if (StringUtils.isEmpty(ref)) {
            return;
        }
        // Hyperlink has 2 case:
        // case 1，In the 'location' tag
        String location = attributes.getValue(Constants.ATTRIBUTE_LOCATION);
        if (location != null) {
            Hyperlink hyperlink = new Hyperlink(location, null, null);
            xssfPreData.hyperlinkMapping.put(new CellAddress(ref), hyperlink);
            return;
        }
        // case 2, In the 'r:id' tag, Then go to 'PackageRelationshipCollection' to get inside
        String rId = attributes.getValue(Constants.ATTRIBUTE_RID);
        if (rId == null || xssfReadContext.packageRelationshipCollection == null) {
            return;
        }
        Optional.ofNullable(xssfReadContext.packageRelationshipCollection.getRelationshipByID(rId))
                .map(PackageRelationship::getTargetURI)
                .ifPresent(uri -> {
                    Hyperlink hyperlink = new Hyperlink(uri.toString(), null, null);
                    xssfPreData.hyperlinkMapping.put(new CellAddress(ref), hyperlink);
                });
    }

    public XSSFPreData getXssfPreData() {
        return xssfPreData;
    }

    public static class XSSFPreData {
        public Map<CellAddress, CellAddress> mergeCellMapping = new HashMap<>();

        public Map<CellAddress, Hyperlink> hyperlinkMapping = new HashMap<>();
    }
}
