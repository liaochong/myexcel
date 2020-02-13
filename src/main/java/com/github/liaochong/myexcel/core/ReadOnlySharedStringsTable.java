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
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.ArrayList;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * @author liaochong
 * @version 1.0
 */
public class ReadOnlySharedStringsTable extends DefaultHandler implements SharedStrings {

    protected final boolean includePhoneticRuns;

    /**
     * An integer representing the total count of strings in the workbook. This count does not
     * include any numbers, it counts only the total of text strings in the workbook.
     */
    protected int count;

    /**
     * An integer representing the total count of unique strings in the Shared String Table.
     * A string is unique even if it is a copy of another string, but has different formatting applied
     * at the character level.
     */
    protected int uniqueCount;

    /**
     * The shared strings table.
     */
    private StringsCache stringsCache;

    private int stringIndex;

    private final boolean rapidity;

    private String[] strings;

    /**
     * Calls {{@link #ReadOnlySharedStringsTable(OPCPackage, boolean, StringsCache, boolean)}} with
     * a value of <code>true</code> for including phonetic runs
     *
     * @param pkg          The {@link OPCPackage} to use as basis for the shared-strings table.
     * @param stringsCache stringsCache
     * @param rapidity     rapidity
     * @throws IOException  If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    public ReadOnlySharedStringsTable(OPCPackage pkg, StringsCache stringsCache, boolean rapidity)
            throws IOException, SAXException {
        this(pkg, true, stringsCache, rapidity);
    }

    /**
     * @param pkg                 The {@link OPCPackage} to use as basis for the shared-strings table.
     * @param includePhoneticRuns whether or not to concatenate phoneticRuns onto the shared string
     * @param stringsCache        stringsCache
     * @param rapidity            rapidity
     * @throws IOException  If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     * @since POI 3.14-Beta3
     */
    public ReadOnlySharedStringsTable(OPCPackage pkg, boolean includePhoneticRuns, StringsCache stringsCache, boolean rapidity)
            throws IOException, SAXException {
        this.includePhoneticRuns = includePhoneticRuns;
        this.stringsCache = stringsCache;
        this.rapidity = rapidity;
        ArrayList<PackagePart> parts =
                pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());

        // Some workbooks have no shared strings table.
        if (parts.size() > 0) {
            PackagePart sstPart = parts.get(0);
            readFrom(sstPart.getInputStream());
        }
    }

    /**
     * Read this shared strings table from an XML file.
     *
     * @param is The input stream containing the XML document.
     * @throws IOException  if an error occurs while reading.
     * @throws SAXException if parsing the XML data fails.
     */
    public void readFrom(InputStream is) throws IOException, SAXException {
        // test if the file is empty, otherwise parse it
        PushbackInputStream pis = new PushbackInputStream(is, 1);
        int emptyTest = pis.read();
        if (emptyTest > -1) {
            pis.unread(emptyTest);
            InputSource sheetSource = new InputSource(pis);
            try {
                XMLReader sheetParser = SAXHelper.newXMLReader();
                sheetParser.setContentHandler(this);
                sheetParser.parse(sheetSource);
                stringsCache.finished();
            } catch (ParserConfigurationException e) {
                throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
            }
        }
    }

    /**
     * Return an integer representing the total count of strings in the workbook. This count does not
     * include any numbers, it counts only the total of text strings in the workbook.
     *
     * @return the total count of strings in the workbook
     */
    @Override
    public int getCount() {
        return this.count;
    }

    /**
     * Returns an integer representing the total count of unique strings in the Shared String Table.
     * A string is unique even if it is a copy of another string, but has different formatting applied
     * at the character level.
     *
     * @return the total count of unique strings in the workbook
     */
    @Override
    public int getUniqueCount() {
        return this.uniqueCount;
    }

    @Override
    public RichTextString getItemAt(int idx) {
        return new XSSFRichTextString(rapidity ? strings[idx] : stringsCache.get(idx));
    }

    //// ContentHandler methods ////

    private StringBuilder characters;
    private boolean tIsOpen;
    private boolean inRPh;

    @Override
    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {
        if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        if ("sst".equals(localName)) {
            String count = attributes.getValue("count");
            if (count != null) {
                this.count = Integer.parseInt(count);
            }
            String uniqueCount = attributes.getValue("uniqueCount");
            if (uniqueCount != null) {
                this.uniqueCount = Integer.parseInt(uniqueCount);
            }
            characters = new StringBuilder(64);
            if (rapidity) {
                this.strings = new String[this.count];
            } else {
                stringsCache.init(this.count);
            }
        } else if ("si".equals(localName)) {
            characters.setLength(0);
        } else if ("t".equals(localName)) {
            tIsOpen = true;
        } else if ("rPh".equals(localName)) {
            inRPh = true;
            //append space...this assumes that rPh always comes after regular <t>
            if (includePhoneticRuns && characters.length() > 0) {
                characters.append(" ");
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {
        if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        if ("si".equals(localName)) {
            if (rapidity) {
                strings[stringIndex++] = characters.toString();
            } else {
                stringsCache.cache(stringIndex++, characters.toString());
            }
        } else if ("t".equals(localName)) {
            tIsOpen = false;
        } else if ("rPh".equals(localName)) {
            inRPh = false;
        }
    }

    /**
     * Captures characters only if a t(ext) element is open.
     */
    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        if (tIsOpen) {
            if (inRPh && includePhoneticRuns) {
                characters.append(ch, start, length);
            } else if (!inRPh) {
                characters.append(ch, start, length);
            }
        }
    }
}
