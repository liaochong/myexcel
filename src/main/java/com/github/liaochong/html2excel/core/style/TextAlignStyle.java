package com.github.liaochong.html2excel.core.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public class TextAlignStyle {

    private static final String LEFT = "left";

    private static final String RIGHT = "right";

    private static final String CENTER = "center";

    private static final String JUSTIFY = "justify";

    public static void setTextAlign(CellStyle cellStyle, Map<String, String> tdStyle) {
        if (Objects.isNull(tdStyle)) {
            return;
        }
        String textAlign = tdStyle.get("text-align");
        if(Objects.isNull(textAlign)){
            return;
        }
        switch (textAlign) {
            case LEFT:
                cellStyle.setAlignment(HorizontalAlignment.LEFT);
                break;
            case RIGHT:
                cellStyle.setAlignment(HorizontalAlignment.RIGHT);
                break;
            case CENTER:
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                break;
            case JUSTIFY:
                cellStyle.setAlignment(HorizontalAlignment.JUSTIFY);
                break;
            default:
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
        }
    }
}
