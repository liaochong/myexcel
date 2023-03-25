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
import org.apache.poi.ss.util.CellAddress;
import org.slf4j.Logger;

import java.util.Collections;
import java.util.List;

/**
 * sax处理
 *
 * @author liaochong
 * @version 1.0
 */
class XSSFSaxReadHandler<T> extends AbstractReadHandler<T> implements XSSFSheetXMLHandler.SheetContentsHandler {

    private static final Logger log = org.slf4j.LoggerFactory.getLogger(XSSFSaxReadHandler.class);
    private int count;
    private final XSSFSheetPreXMLHandler.XSSFPreData xssfPreData;

    public XSSFSaxReadHandler(
            List<T> result,
            SaxExcelReader.ReadConfig<T> readConfig, XSSFSheetPreXMLHandler.XSSFPreData xssfPreData) {
        super(false, result, readConfig, xssfPreData != null ? xssfPreData.mergeCellMapping : Collections.emptyMap());
        this.xssfPreData = xssfPreData;
    }

    @Override
    public void startRow(int rowNum, boolean newInstance) {
        newRow(rowNum, newInstance);
    }

    @Override
    public void endRow(int rowNum) {
        handleResult();
        count++;
    }

    @Override
    public void cell(CellAddress cellAddress, String formattedValue) {
        if (xssfPreData != null) {
            Hyperlink hyperlink = xssfPreData.hyperlinkMapping.get(cellAddress);
            if (hyperlink != null) {
                hyperlink.setLabel(formattedValue);
            }
            this.readContext.setHyperlink(hyperlink);
        }
        int thisCol = cellAddress.getColumn();
        handleField(thisCol, formattedValue);
    }

    @Override
    public void endSheet() {
        log.info("Import completed, total number of rows {}.", count);
    }
}
