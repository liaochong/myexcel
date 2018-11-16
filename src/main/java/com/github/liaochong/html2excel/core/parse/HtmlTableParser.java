package com.github.liaochong.html2excel.core.parse;

import com.github.liaochong.html2excel.utils.StyleUtils;
import com.github.liaochong.html2excel.utils.TdUtils;
import org.apache.commons.codec.CharEncoding;
import org.apache.commons.collections4.CollectionUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.IOException;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Predicate;
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
        return IntStream.range(0, tableElements.size()).mapToObj(i -> {
            Element tableElement = tableElements.get(i);
            Table table = new Table();
            table.setIndex(i);
            table.setTableElement(tableElement);

            Elements captionElements = tableElement.getElementsByTag(TableTag.caption.name());
            if (!captionElements.isEmpty()) {
                table.setCaption(captionElements.next().text());
            }

            Map<String, String> tableStyleMap = StyleUtils.parseStyle(tableElement);
            table.setStyleMap(tableStyleMap);

            this.parseTrOfTable(tableElement, table, tableStyleMap);
            return table;
        }).collect(Collectors.toList());
    }

    /**
     * 解析table中的tr
     *
     * @param tableElement  table元素
     * @param table         table
     * @param tableStyleMap table style map
     */
    private void parseTrOfTable(Element tableElement, Table table, Map<String, String> tableStyleMap) {
        Elements trElements = tableElement.getElementsByTag(TableTag.tr.name());
        if (trElements.isEmpty()) {
            return;
        }
        Map<Element, Map<String, String>> upperStyleMap = new HashMap<>();

        List<Tr> trList = IntStream.range(0, trElements.size()).parallel().mapToObj(index -> {
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
            tr.setTrElement(trElement);
            tr.setStyle(StyleUtils.mixStyle(upperStyle, StyleUtils.parseStyle(trElement)));
            this.parseTdOfTr(trElement, tr, table);
            return tr;
        }).collect(Collectors.toList());

        table.setTrList(trList);

        int lastColumnNum = trList.parallelStream().max(Comparator.comparing(Tr::getLastColumnNum)).get().getLastColumnNum();
        table.setLastColumnNum(lastColumnNum);
        // 调整td位置
        // 排除第一行，第一行不需要进行调整
        if (trList.size() > 1) {
            trList.subList(1, trList.size()).parallelStream().forEach(tr -> {
                tr.getTdList().parallelStream().forEach(td -> this.adjustTdPosition(td, tr.getIndex(), trList, lastColumnNum));
            });
        }
    }

    /**
     * 获取tr中的td
     *
     * @param trElement tr元素
     * @param tr        tr
     * @param table     table
     */
    private void parseTdOfTr(Element trElement, Tr tr, Table table) {
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
                int shift = tr.getTdList().stream().filter(t -> t.getColSpan() > 0)
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
            tr.getTdList().add(td);

            // 设置每列最宽宽度
            int width = TdUtils.getStringWidth(td.getContent());
            Integer maxWidth = table.getColMaxWidthMap().get(td.getCol());
            if (Objects.isNull(maxWidth) || maxWidth < width) {
                table.getColMaxWidthMap().put(td.getCol(), width);
            }

            int colIndex = TdUtils.get(td::getColSpan, td::getCol);
            if (Objects.isNull(tr.getLastColumnNum()) || colIndex > tr.getLastColumnNum()) {
                tr.setLastColumnNum(colIndex);
            }
        }
    }

    /**
     * 调整表格单元格位置
     *
     * @param td      单元格
     * @param trIndex 单元格所在行索引
     */
    private void adjustTdPosition(Td td, int trIndex, List<Tr> trList, int lastColumnNum) {
        Predicate<Tr> predicate = tr -> tr.getIndex() < trIndex;
        List<Td> tds = trList.stream().filter(predicate).flatMap(tr -> tr.getTdList().stream())
                .collect(Collectors.toList());
        if (CollectionUtils.isEmpty(tds)) {
            return;
        }
        for (int i = 0; i < lastColumnNum; i++) {
            Td td1 = tds.stream().filter(prevTd -> prevTd.getCol() <= td.getCol()
                    && TdUtils.get(prevTd::getRowSpan, prevTd::getRow) >= td.getRow()).findFirst().orElse(null);
            if (Objects.isNull(td1)) {
                return;
            }
            int prevTdColSpan = td1.getColSpan();
            int realCol = prevTdColSpan > 0 ? td.getCol() + prevTdColSpan : td.getCol() + 1;
            td.setCol(realCol);
            tds.remove(td1);
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
