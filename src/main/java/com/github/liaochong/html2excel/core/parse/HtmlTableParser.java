package com.github.liaochong.html2excel.core.parse;

import com.github.liaochong.html2excel.utils.StyleUtils;
import com.github.liaochong.html2excel.utils.TdUtils;
import org.apache.commons.codec.CharEncoding;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.IOException;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author liaochong
 * @version 1.0
 */
public class HtmlTableParser {

    /**
     * html解析后文档
     */
    private Document document;

    private HtmlTableParser() {

    }

    public static HtmlTableParser of(File htmlFile) throws IOException {
        Objects.requireNonNull(htmlFile);
        HtmlTableParser parser = new HtmlTableParser();
        parser.document = Jsoup.parse(htmlFile, CharEncoding.UTF_8);
        return parser;
    }

    /**
     * 获取所有表格
     *
     * @return 所有表格
     */
    public List<Table> getAllTable() {
        Elements tableElements = document.getElementsByTag(TableTag.table.name());
        List<Table> tableList = tableElements.stream().map(tableElement -> {
            Table table = new Table();
            Map<String, String> tableStyleMap = StyleUtils.parseStyle(tableElement);
            List<Tr> trList = this.getTrOfTable(tableElement, table, tableStyleMap);
            table.setTrList(trList);
            table.setStyleMap(tableStyleMap);
            return table;
        }).collect(Collectors.toList());
        // 设置表格最后列
        tableList.forEach(table -> {
            Integer lastColumnNum = table.getTrList().parallelStream().mapToInt(Tr::getLastColumnNum).max().orElseGet(() -> 0);
            table.setLastColumnNum(lastColumnNum);
        });
        return tableList;
    }

    /**
     * @param tableElement
     * @param table
     * @param tableStyleMap
     * @return
     */
    private List<Tr> getTrOfTable(Element tableElement, Table table, Map<String, String> tableStyleMap) {
        Elements trElements = tableElement.getElementsByTag(TableTag.tr.name());
        if (trElements.isEmpty()) {
            return Collections.emptyList();
        }
        Map<Element, Map<String, String>> upperStyleMap = new HashMap<>();

        return IntStream.range(0, trElements.size()).parallel().mapToObj(index -> {
            Element trElement = trElements.get(index);
            Element parent = trElement.parent();
            Map<String, String> upperStyle;
            if (upperStyleMap.containsKey(parent)) {
                upperStyle = upperStyleMap.get(parent);
            } else {
                upperStyle = StyleUtils.mixStyle(tableStyleMap, StyleUtils.parseStyle(parent));
                upperStyleMap.put(parent, upperStyle);
            }
            Tr tr = new Tr(index);
            tr.setStyle(StyleUtils.mixStyle(upperStyle, StyleUtils.parseStyle(trElement)));
            this.getTdOfTr(trElement, tr, table);
            return tr;
        }).collect(Collectors.toList());
    }

    /**
     * 获取tr中的td
     *
     * @param trElement tr元素
     * @param tr        tr
     * @param table     table
     */
    private void getTdOfTr(Element trElement, Tr tr, Table table) {
        Elements childrenElements = trElement.children();
        if (childrenElements.isEmpty()) {
            return;
        }
        for (int i = 0, size = childrenElements.size(); i < size; i++) {
            Element tdElement = childrenElements.get(i);
            Td td = new Td();
            td.setTh(Objects.equals(TableTag.th.name(), tdElement.tagName()));
            td.setRow(tr.getIndex());
            td.setStyle(StyleUtils.mixStyle(tr.getStyle(), StyleUtils.parseStyle(tdElement)));
            // 除每行第一个单元格外，修正含跨列的单元格位置
            if (i > 0) {
                int shift = tr.getTds().stream().filter(t -> t.getColSpan() > 0)
                        .mapToInt(t -> t.getColSpan() - 1).sum();
                td.setCol(i + shift);
            } else {
                td.setCol(i);
            }

            String colSpan = tdElement.attr(TableTag.colspan.name());
            td.setColSpan(TdUtils.getSpan(colSpan));

            String rowSpan = tdElement.attr(TableTag.rowspan.name());
            td.setRowSpan(TdUtils.getSpan(rowSpan));

            td.setContent(tdElement.text());
            tr.getTds().add(td);

            // 设置每列最宽宽度
            int width = TdUtils.getStringWidth(td.getContent());
            Integer maxWidth = table.getColMaxWidthMap().get(td.getCol());
            if (Objects.isNull(maxWidth) || maxWidth < width) {
                table.getColMaxWidthMap().put(td.getCol(), width);
            }

            int colIndex = TdUtils.get(td::getColSpan, td::getCol);
            if (colIndex > tr.getLastColumnNum()) {
                tr.setLastColumnNum(colIndex);
            }
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
