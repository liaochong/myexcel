package com.github.liaochong.html2excel.core.parse;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;
import org.jsoup.nodes.Element;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class Table {

    Element tableElement;

    String caption;

    List<Tr> trList;

    Map<String, String> styleMap;

    Integer lastColumnNum;

    Integer lastRowNum;

    Integer index;

    Map<Integer, Integer> colMaxWidthMap = new HashMap<>();
}
