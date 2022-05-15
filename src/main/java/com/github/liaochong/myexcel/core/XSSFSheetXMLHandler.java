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
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * @author liaochong
 * @version 1.0
 */
class XSSFSheetXMLHandler extends DefaultHandler {
    private static final Logger logger = LoggerFactory.getLogger(XSSFSheetXMLHandler.class);

    /**
     * These are the different kinds of cells we support.
     * We keep track of the current one between
     * the start and end.
     */
    enum xssfDataType {
        BOOLEAN,
        ERROR,
        FORMULA,
        INLINE_STRING,
        SST_STRING,
        NUMBER,
    }

    /**
     * Read only access to the shared strings table, for looking
     * up (most) string cell's contents
     */
    private final SharedStrings sharedStringsTable;

    /**
     * Where our text is going
     */
    private final XSSFSheetXMLHandler.SheetContentsHandler output;

    // Set when V start element is seen
    private boolean vIsOpen;
    // Set when an Inline String "is" is seen
    private boolean isIsOpen;

    // Set when cell start element is seen;
    // used when cell close element is seen.
    private xssfDataType nextDataType;

    private int rowNum;
    private int preRowNum = -1;
    private int nextRowNum;      // some sheets do not have rowNums, Excel can read them so we should try to handle them correctly as well
    private String cellRef;

    private final boolean detectedMerge;
    private long waitCount = 1;

    // Gathers characters as they are seen.
    private final StringBuilder value = new StringBuilder(64);

    private final Map<CellAddress, CellAddress> mergeCellMapping;

    private final Map<CellAddress, String> mergeFirstCellMapping;

    /**
     * Accepts objects needed while parsing.
     *
     * @param strings              Table of shared strings
     * @param sheetContentsHandler sheetContentsHandler
     */
    public XSSFSheetXMLHandler(
            Map<CellAddress, CellAddress> mergeCellMapping,
            SharedStrings strings,
            XSSFSheetXMLHandler.SheetContentsHandler sheetContentsHandler) {
        this.mergeCellMapping = mergeCellMapping;
        this.detectedMerge = !mergeCellMapping.isEmpty();
        this.mergeFirstCellMapping = mergeCellMapping.values().stream().distinct().collect(Collectors.toMap(cellAddress -> cellAddress, c -> ""));
        this.sharedStringsTable = strings;
        this.output = sheetContentsHandler;
        this.nextDataType = xssfDataType.NUMBER;
    }

    private boolean isTextTag(String name) {
        if ("v".equals(name)) {
            // Easy, normal v text tag
            return true;
        }
        if ("inlineStr".equals(name)) {
            // Easy inline string
            return true;
        }
        // Inline string <is><t>...</t></is> pair
        return "t".equals(name) && isIsOpen;
        // It isn't a text tag
    }

    @Override
    @SuppressWarnings("unused")
    public void startElement(String uri, String localName, String qName,
                             Attributes attributes) throws SAXException {

        if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        if (isTextTag(localName)) {
            vIsOpen = true;
            // Clear contents cache
            value.setLength(0);
        }
        // c => cell
        else if ("c".equals(localName)) {
            // Set up defaults.
            this.nextDataType = xssfDataType.NUMBER;
            cellRef = attributes.getValue("r");
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("b".equals(cellType))
                nextDataType = xssfDataType.BOOLEAN;
            else if ("e".equals(cellType))
                nextDataType = xssfDataType.ERROR;
            else if ("inlineStr".equals(cellType))
                nextDataType = xssfDataType.INLINE_STRING;
            else if ("s".equals(cellType))
                nextDataType = xssfDataType.SST_STRING;
            else if ("str".equals(cellType))
                nextDataType = xssfDataType.FORMULA;
        } else if ("row".equals(localName)) {
            String rowNumStr = attributes.getValue("r");
            if (rowNumStr != null) {
                rowNum = Integer.parseInt(rowNumStr) - 1;
            } else {
                rowNum = nextRowNum;
            }
            if (rowNum - 1 != preRowNum) {
                for (int blankRowNum = preRowNum + 1; blankRowNum < rowNum; blankRowNum++) {
                    output.startRow(blankRowNum, true);
                    output.endRow(blankRowNum);
                }
            }
            output.startRow(rowNum, !detectedMerge || --waitCount == 0);
            if (detectedMerge && waitCount == 0) {
                waitCount = mergeCellMapping.values().stream().filter(c -> Objects.equals(c.getRow(), rowNum)).count() + 1;
            }
            this.preRowNum = rowNum;
        } else if ("is".equals(localName)) {
            // Inline string outer tag
            isIsOpen = true;
        } else if ("f".equals(localName)) {
            // Mark us as being a formula if not already
            if (nextDataType == xssfDataType.NUMBER) {
                nextDataType = xssfDataType.FORMULA;
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName)
            throws SAXException {

        if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        String thisStr = null;

        // v => contents of a cell
        if (isTextTag(localName)) {
            vIsOpen = false;

            // Process the value contents as required, now we have it all
            switch (nextDataType) {
                case BOOLEAN:
                    char first = value.charAt(0);
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    break;

                case ERROR:
                    thisStr = "ERROR:" + value;
                    break;

                case FORMULA:
                    // No formatting applied, just do raw value in all cases
                    thisStr = value.toString();
                    break;
                case INLINE_STRING:
                    // TODO: Can these ever have formatting on them?
                    XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                    thisStr = rtsi.toString();
                    break;

                case SST_STRING:
                    String sstIndex = value.toString();
                    try {
                        int idx = Integer.parseInt(sstIndex);
                        RichTextString rtss = sharedStringsTable.getItemAt(idx);
                        thisStr = rtss.toString();
                    } catch (NumberFormatException ex) {
                        logger.error("Failed to parse SST index '" + sstIndex, ex);
                    }
                    break;

                case NUMBER:
                    String n = value.toString();
                    if (n.contains(Constants.SPOT)) {
                        n = String.valueOf(Double.parseDouble(n));
                    }
                    thisStr = n;
                    break;

                default:
                    thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                    break;
            }

            CellAddress cellAddress = new CellAddress(cellRef);
            String finalThisStr = thisStr;
            mergeFirstCellMapping.computeIfPresent(cellAddress, (k, v) -> finalThisStr);
            // Output
            output.cell(cellAddress, thisStr);
        } else if ("c".equals(localName)) {
            CellAddress cellAddress = new CellAddress(cellRef);
            CellAddress firstCellAddress = mergeCellMapping.get(cellAddress);
            if (firstCellAddress != null) {
                output.cell(cellAddress, mergeFirstCellMapping.get(firstCellAddress));
                mergeCellMapping.remove(cellAddress);
            }
        } else if ("row".equals(localName)) {
            // Finish up the row
            if (!detectedMerge || waitCount == 0) {
                output.endRow(rowNum);
            }
            // some sheets do not have rowNum set in the XML, Excel can read them so we should try to read them as well
            nextRowNum = rowNum + 1;
        } else if ("sheetData".equals(localName)) {
            // indicate that this sheet is now done
            output.endSheet();
        } else if ("is".equals(localName)) {
            isIsOpen = false;
        }
    }

    /**
     * Captures characters only if a suitable element is open.
     * Originally was just "v"; extended for inlineStr also.
     */
    @Override
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        if (vIsOpen) {
            value.append(ch, start, length);
        }
    }

    /**
     * You need to implement this to handle the results
     * of the sheet parsing.
     */
    public interface SheetContentsHandler {
        /**
         * A row with the (zero based) row number has started
         *
         * @param rowNum      rowNum
         * @param newInstance newInstance
         */
        void startRow(int rowNum, boolean newInstance);

        /**
         * A row with the (zero based) row number has ended
         *
         * @param rowNum rowNum
         */
        void endRow(int rowNum);

        /**
         * A cell, with the given formatted value (may be null),
         * and possibly a comment (may be null), was encountered.
         * <p>
         * Sheets that have missing or empty cells may result in
         * sparse calls to <code>cell</code>. See the code in
         * <code>src/examples/src/org/apache/poi/xssf/eventusermodel/XLSX2CSV.java</code>
         * for an example of how to handle this scenario.
         *
         * @param cellAddress    cellAddress
         * @param formattedValue formattedValue
         */
        void cell(CellAddress cellAddress, String formattedValue);

        /**
         * Signal that the end of a sheet was been reached
         */
        default void endSheet() {
        }
    }
}
