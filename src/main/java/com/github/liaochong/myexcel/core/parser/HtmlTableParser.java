/*
 * Copyright 2017 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core.parser;

import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.style.FontStyle;
import com.github.liaochong.myexcel.utils.ImageUtil;
import com.github.liaochong.myexcel.utils.RegexpUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.StyleUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * html table parser
 *
 * @author liaochong
 * @version 1.0
 */
public class HtmlTableParser {

    private static final Pattern DOUBLE_PATTERN = Pattern.compile("^[-+]?(\\d+(\\.\\d*)?|\\.\\d+)([eE]([-+]?([012]?\\d{1,2}|30[0-7])|-3([01]?[4-9]|[012]?[0-3])))?[dD]?$");

    private static final Pattern LINE_FEED_PATTERN = Pattern.compile("\\\\n");
    private static final Logger log = org.slf4j.LoggerFactory.getLogger(HtmlTableParser.class);

    private ParseConfig parseConfig;

    private File htmlFile;

    private String html;

    private final Map<String, String> defaultLinkStyle = new HashMap<>();

    private XSSFRichTextString spanText;

    private HtmlTableParser() {
        defaultLinkStyle.put(FontStyle.FONT_COLOR, "blue");
        defaultLinkStyle.put(FontStyle.TEXT_DECORATION, FontStyle.UNDERLINE);
    }

    public static HtmlTableParser of(File htmlFile) {
        Objects.requireNonNull(htmlFile);
        HtmlTableParser htmlTableParser = new HtmlTableParser();
        htmlTableParser.htmlFile = htmlFile;
        return htmlTableParser;
    }

    public static HtmlTableParser of(String html) {
        Objects.requireNonNull(html);
        HtmlTableParser htmlTableParser = new HtmlTableParser();
        htmlTableParser.html = html;
        return htmlTableParser;
    }

    /**
     * 获取所有表格
     *
     * @param parseConfig 解析配置
     * @return 所有表格
     * @throws IOException IOException
     */
    public List<Table> getAllTable(ParseConfig parseConfig) throws IOException {
        log.info("Start parsing html file");
        long startTime = System.currentTimeMillis();
        Document document;
        if (htmlFile != null) {
            document = Jsoup.parse(htmlFile, CharEncoding.UTF_8);
        } else {
            document = Jsoup.parse(html, CharEncoding.UTF_8);
        }
        document.outputSettings(new Document.OutputSettings().prettyPrint(false));
        //select all <br> tags and append \n after that
        document.select("br").after("\\n");
        //select all <p> tags and prepend \n before that
        document.select("p").before("\\n");
        this.parseConfig = parseConfig;
        Elements tableElements = document.getElementsByTag(HtmlTag.table.name());
        List<Table> result = tableElements.stream().map(tableElement -> {
            Table table = new Table();
            Elements captionElements = tableElement.getElementsByTag(HtmlTag.caption.name());
            if (!captionElements.isEmpty()) {
                table.caption = captionElements.first().text();
            }
            this.parseTrOfTable(table, tableElement, StyleUtil.parseStyle(tableElement));
            return table;
        }).collect(Collectors.toList());
        log.info("Complete html file parsing,takes {} ms", System.currentTimeMillis() - startTime);
        return result;
    }

    /**
     * 解析table中的tr
     *
     * @param table table
     */
    private void parseTrOfTable(Table table, Element tableElement, Map<String, String> tableStyle) {
        Map<Element, Map<String, String>> parentStyleMap = new ConcurrentHashMap<>();

        Elements trElements = tableElement.getElementsByTag(HtmlTag.tr.name());
        final Map<Integer, List<Integer>> seizeMap = new HashMap<>();
        List<Tr> trList = IntStream.range(0, trElements.size()).mapToObj(index -> {
            Element trElement = trElements.get(index);
            Element parent = trElement.parent();
            Map<String, String> upperStyle;
            if (Objects.equals(parent, tableElement)) {
                upperStyle = tableStyle;
            } else {
                if (parentStyleMap.containsKey(parent)) {
                    upperStyle = parentStyleMap.get(parent);
                } else {
                    upperStyle = StyleUtil.mixStyle(tableStyle, StyleUtil.parseStyle(parent));
                    parentStyleMap.putIfAbsent(parent, upperStyle);
                }
            }
            Map<String, String> trStyleMap = StyleUtil.mixStyle(upperStyle, StyleUtil.parseStyle(trElement));
            String height = trStyleMap.get("height");
            Tr tr = new Tr(index, TdUtil.getValue(height), true);
            // 行可见性
            tr.visibility = !Objects.equals(trStyleMap.get("visibility"), "hidden");
            this.parseTdOfTr(tr, trElement, trStyleMap, seizeMap);
            return tr;
        }).collect(Collectors.toCollection(LinkedList::new));
        table.trList = trList;
    }

    /**
     * 获取tr中的td
     *
     * @param tr        tr
     * @param trElement trElement
     * @param trStyle   trStyle
     * @param seizeMap  seizeMap 占位map
     */
    private void parseTdOfTr(Tr tr, Element trElement, Map<String, String> trStyle, Map<Integer, List<Integer>> seizeMap) {
        Elements tdElements = trElement.children();
        if (tdElements.isEmpty()) {
            tr.tdList = Collections.emptyList();
            tr.colWidthMap = Collections.emptyMap();
            return;
        }

        final List<Td> tdList = new LinkedList<>();
        final Map<Integer, Integer> colWidthMap = new HashMap<>(tdElements.size());
        List<Integer> seizeOfTr = seizeMap.getOrDefault(tr.index, Collections.emptyList());
        // 单元格偏移量
        int shift = 0;
        for (int i = 0, size = tdElements.size(); i < size; i++) {
            Element tdElement = tdElements.get(i);
            Td td = new Td(tr.index, i + shift);
            this.setTdContent(tdElement, td);

            td.th = Objects.equals(HtmlTag.th.name(), tdElement.tagName());
            Map<String, String> tdStyle = StyleUtil.parseStyle(tdElement);
            if (tdStyle.isEmpty() && ContentTypeEnum.isLink(td.tdContentType)) {
                tdStyle = defaultLinkStyle;
            }
            td.style = StyleUtil.mixStyle(trStyle, tdStyle);

            String colSpan = tdElement.attr(HtmlTag.colspan.name());
            td.setColSpan(TdUtil.getSpan(colSpan));

            String rowSpan = tdElement.attr(HtmlTag.rowspan.name());
            td.setRowSpan(TdUtil.getSpan(rowSpan));

            if (!seizeOfTr.isEmpty()) {
                List<Integer> checkedPositions = new ArrayList<>();
                while (true) {
                    List<Integer> seizePositions = seizeOfTr.stream().filter(s -> td.col >= s).collect(Collectors.toList());
                    if (!checkedPositions.isEmpty()) {
                        seizePositions.removeAll(checkedPositions);
                    }
                    if (seizePositions.isEmpty()) {
                        break;
                    }
                    td.col = td.col + seizePositions.size();
                    checkedPositions.addAll(seizePositions);
                }
            }

            if (td.rowSpan > 1) {
                for (int j = 1, length = td.rowSpan; j < length; j++) {
                    int rowNum = tr.index + j;
                    List<Integer> seizePosOfTr = seizeMap.get(rowNum);
                    if (Objects.isNull(seizePosOfTr)) {
                        seizePosOfTr = new ArrayList<>();
                        seizeMap.put(rowNum, seizePosOfTr);
                    }
                    IntStream.rangeClosed(td.col, td.getColBound()).forEach(seizePosOfTr::add);
                }
            }

            if (td.colSpan > 0) {
                shift += td.colSpan - 1;
            }
            tdList.add(td);
            // 斜线
            this.setSlant(tdElement, td);
            // 设置每列宽度
            this.setColumnWidth(colWidthMap, td);
            // 批注
            this.setComment(tdElement, td);
        }
        tr.tdList = tdList;
        tr.colWidthMap = colWidthMap;
    }

    private void setComment(Element tdElement, Td td) {
        String commentText = tdElement.attr("comment-text");
        String author = tdElement.attr("comment-author");
        if (StringUtil.isBlank(commentText) && StringUtil.isBlank(author)) {
            return;
        }
        Comment comment = new Comment();
        comment.text = commentText;
        comment.author = author;
        td.comment = comment;
    }

    private void setSlant(Element tdElement, Td td) {
        boolean hasSlant = tdElement.hasAttr("slant");
        if (hasSlant) {
            String slantStr = tdElement.attr("slant");
            if (StringUtil.isNotBlank(slantStr)) {
                String[] splits = slantStr.split(" ");
                if (splits.length != 3) {
                    throw new IllegalArgumentException("Slash setting error");
                }
                td.slant = new Slant(LineStyleEnum.getByName(splits[0]), splits[1], splits[2]);
            } else {
                td.slant = new Slant();
            }
        }
    }

    private void setColumnWidth(Map<Integer, Integer> colWidthMap, Td td) {
        if (parseConfig.isComputeAutoWidth()) {
            int width = TdUtil.getStringWidth(td.content);
            if (td.colSpan > 1) {
                int realWidth = (int) Math.ceil(width * 1.0 / td.colSpan);
                for (int j = 0, span = td.colSpan; j < span; j++) {
                    int colIndex = td.col + j;
                    Integer colWidth = colWidthMap.get(colIndex);
                    if (colWidth == null || colWidth < realWidth) {
                        colWidthMap.put(colIndex, realWidth);
                    }
                }
            } else {
                colWidthMap.put(td.col, width);
            }
        }
        String widthStr = td.style.get("width");
        if (widthStr != null) {
            int width = TdUtil.getValue(widthStr);
            if (width >= 0) {
                colWidthMap.put(td.col, width);
            }
        }
    }

    private void setTdContent(Element tdElement, Td td) {
        Elements imgs = tdElement.getElementsByTag(HtmlTag.img.name());
        if (!imgs.isEmpty()) {
            String src = imgs.get(0).attr("src");
            if (src.startsWith(Constants.HTTP)) {
                td.fileIs = ImageUtil.getImageFromNetByUrl(src);
            } else {
                td.file = new File(src);
            }
            td.tdContentType = ContentTypeEnum.IMAGE;
            return;
        }
        Elements links = tdElement.getElementsByTag(HtmlTag.a.name());
        if (!links.isEmpty()) {
            Element a = links.get(0);
            td.content = a.text();
            String href = a.attr("href").trim();
            td.link = href;
            td.tdContentType = href.startsWith("mailto:") ? ContentTypeEnum.LINK_EMAIL : ContentTypeEnum.LINK_URL;
            return;
        }
        String content = this.parseContent(tdElement, td);
        td.content = content;
        if (StringUtil.isBlank(content)) {
            return;
        }
        if (tdElement.hasAttr("string")) {
            return;
        }
        if (tdElement.hasAttr("double")) {
            td.tdContentType = ContentTypeEnum.DOUBLE;
            td.content = RegexpUtil.removeComma(td.content);
            return;
        }
        // 公式设置
        boolean isFormula = tdElement.hasAttr("formula");
        if (isFormula) {
            td.formula = true;
            String formula = td.content.trim();
            if (formula.startsWith(Constants.EQUAL)) {
                formula = formula.substring(1);
            }
            td.content = formula;
            return;
        }
        if (tdElement.hasAttr("url")) {
            String link = tdElement.attr("url");
            td.tdContentType = ContentTypeEnum.LINK_URL;
            td.link = link;
            return;
        }
        if (tdElement.hasAttr("email")) {
            String link = tdElement.attr("email");
            td.tdContentType = ContentTypeEnum.LINK_EMAIL;
            td.link = link;
            return;
        }
        if (tdElement.hasAttr("dropDownList")) {
            td.tdContentType = ContentTypeEnum.DROP_DOWN_LIST;
            return;
        }
        if (Constants.TRUE.equals(content) || Constants.FALSE.equals(content)) {
            td.tdContentType = ContentTypeEnum.BOOLEAN;
            return;
        }
        if (DOUBLE_PATTERN.matcher(content).matches()) {
            td.tdContentType = ContentTypeEnum.DOUBLE;
        }
    }

    private String parseContent(Element tdElement, Td td) {
        Elements spans = tdElement.getElementsByTag(HtmlTag.span.name());
        if (!spans.isEmpty()) {
            td.fonts = new LinkedList<>();
            if (spanText == null) {
                spanText = new XSSFRichTextString("");
            }
            int startIndex = 0;
            for (Element spanElement : spans) {
                String spanContent = spanElement.text();
                spanContent = LINE_FEED_PATTERN.matcher(spanContent).replaceAll("\n");
                spanText.setString(spanContent);
                Font font = new Font();
                font.startIndex = startIndex;
                font.endIndex = startIndex + spanText.length();

                Map<String, String> fontStyle = StyleUtil.parseStyle(spanElement);
                if (!fontStyle.isEmpty()) {
                    font.style = fontStyle;
                    td.fonts.add(font);
                }
                startIndex = font.endIndex;
            }
        }
        return LINE_FEED_PATTERN.matcher(tdElement.text()).replaceAll("\n");
    }

    public enum HtmlTag {
        /**
         * table
         */
        table,
        /**
         * caption
         */
        caption,
        /**
         * thead
         */
        thead,
        /**
         * tbody
         */
        tbody,
        /**
         * tr
         */
        tr,
        /**
         * th
         */
        th,
        /**
         * td
         */
        td,
        /**
         * colspan
         */
        colspan,
        /**
         * rowspan
         */
        rowspan,
        /**
         * link
         */
        link,
        /**
         * img
         */
        img,
        /**
         * a
         */
        a,
        /**
         * span
         */
        span
    }
}
