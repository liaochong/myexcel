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

import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.StyleUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.codec.CharEncoding;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

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
@Slf4j
public class HtmlTableParser {

    private static final Pattern DOUBLE_PATTERN = Pattern.compile("^[-+]?(\\d+(\\.\\d*)?|\\.\\d+)([eE]([-+]?([012]?\\d{1,2}|30[0-7])|-3([01]?[4-9]|[012]?[0-3])))?[dD]?$");

    private ParseConfig parseConfig;

    private File htmlFile;

    private String html;

    private HtmlTableParser() {

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
        if (Objects.nonNull(htmlFile)) {
            document = Jsoup.parse(htmlFile, CharEncoding.UTF_8);
        } else {
            document = Jsoup.parse(html, CharEncoding.UTF_8);
        }
        this.parseConfig = parseConfig;
        Elements tableElements = document.getElementsByTag(TableTag.table.name());
        List<Table> result = tableElements.stream().map(tableElement -> {
            Table table = new Table();
            Elements captionElements = tableElement.getElementsByTag(TableTag.caption.name());
            if (!captionElements.isEmpty()) {
                table.setCaption(captionElements.first().text());
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

        Elements trElements = tableElement.getElementsByTag(TableTag.tr.name());
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
            Tr tr = new Tr(index);
            // 行可见性
            Map<String, String> trStyleMap = StyleUtil.mixStyle(upperStyle, StyleUtil.parseStyle(trElement));
            tr.setVisibility(!Objects.equals(trStyleMap.get("visibility"), "hidden"));
            this.parseTdOfTr(tr, trElement, trStyleMap, seizeMap);
            return tr;
        }).collect(Collectors.toList());
        table.setTrList(trList);
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
            tr.setTdList(Collections.emptyList());
            tr.setColWidthMap(Collections.emptyMap());
            return;
        }

        final List<Td> tdList = new LinkedList<>();
        final Map<Integer, Integer> colWidthMap = new HashMap<>(tdElements.size());
        List<Integer> seizeOfTr = seizeMap.getOrDefault(tr.getIndex(), Collections.emptyList());
        // 单元格偏移量
        int shift = 0;
        for (int i = 0, size = tdElements.size(); i < size; i++) {
            Element tdElement = tdElements.get(i);
            Td td = new Td();
            this.setTdContent(tdElement, td);

            td.setTh(Objects.equals(TableTag.th.name(), tdElement.tagName()));
            td.setRow(tr.getIndex());
            td.setStyle(StyleUtil.mixStyle(trStyle, StyleUtil.parseStyle(tdElement)));
            // 除每行第一个单元格外，修正含跨列的单元格位置
            td.setCol(i + shift);

            String colSpan = tdElement.attr(TableTag.colspan.name());
            td.setColSpan(TdUtil.getSpan(colSpan));

            String rowSpan = tdElement.attr(TableTag.rowspan.name());
            td.setRowSpan(TdUtil.getSpan(rowSpan));

            if (!seizeOfTr.isEmpty()) {
                List<Integer> checkedPositions = new ArrayList<>();
                while (true) {
                    List<Integer> seizePositions = seizeOfTr.stream().filter(s -> td.getCol() >= s).collect(Collectors.toList());
                    if (!checkedPositions.isEmpty()) {
                        seizePositions.removeAll(checkedPositions);
                    }
                    if (seizePositions.isEmpty()) {
                        break;
                    }
                    td.setCol(td.getCol() + seizePositions.size());
                    checkedPositions.addAll(seizePositions);
                }
            }

            int rowBound = TdUtil.get(td::getRowSpan, td::getRow);
            td.setRowBound(rowBound);

            int colBound = TdUtil.get(td::getColSpan, td::getCol);
            td.setColBound(colBound);

            if (td.getRowSpan() > 1) {
                for (int j = 1, length = td.getRowSpan(); j < length; j++) {
                    int rowNum = tr.getIndex() + j;
                    List<Integer> seizePosOfTr = seizeMap.get(rowNum);
                    if (Objects.isNull(seizePosOfTr)) {
                        seizePosOfTr = new ArrayList<>();
                        seizeMap.put(rowNum, seizePosOfTr);
                    }
                    IntStream.rangeClosed(td.getCol(), td.getColBound()).forEach(seizePosOfTr::add);
                }
            }

            if (td.getColSpan() > 0) {
                shift += td.getColSpan() - 1;
            }
            tdList.add(td);

            // 设置每列宽度
            if (parseConfig.isComputeAutoWidth()) {
                int width = TdUtil.getStringWidth(td.getContent());
                colWidthMap.put(td.getCol(), width);
            } else if (parseConfig.isCustomWidth()) {
                String widthStr = td.getStyle().get("width");
                if (Objects.nonNull(widthStr)) {
                    Integer width = Integer.valueOf(widthStr.replaceAll("\\D*", ""));
                    if (width > 0) {
                        colWidthMap.put(td.getCol(), width);
                    }
                }
            }
        }
        tr.setTdList(tdList);
        tr.setColWidthMap(colWidthMap);
    }

    private void setTdContent(Element tdElement, Td td) {
        // 公式设置
        td.setFormula(tdElement.hasAttr("formula"));
        String tdContent = tdElement.text();
        td.setContent(tdContent);
        if (StringUtil.isBlank(tdContent)) {
            return;
        }
        if (tdElement.hasAttr("string")) {
            return;
        }
        if (Objects.equals(tdContent, "true") || Objects.equals(tdContent, "false")) {
            td.setTdContentType(ContentTypeEnum.BOOLEAN);
            return;
        }
        if (DOUBLE_PATTERN.matcher(tdContent).matches()) {
            td.setTdContentType(ContentTypeEnum.DOUBLE);
        }
    }

    public enum TableTag {
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
        rowspan;
    }
}
