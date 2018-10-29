package com.github.liaochong.html2excel.core.style;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;

import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public class BackgroundStyle {

    private static final String HASH = "#";

    private static final String RGB = "rgb";

    private static Map<String, HSSFColor.HSSFColorPredefined> colorPredefinedMap;

    static {
        colorPredefinedMap = Arrays.stream(HSSFColor.HSSFColorPredefined.values())
                .collect(Collectors.toMap(c -> c.toString().toLowerCase().replaceAll("_", ""), c -> c));
    }


    public static void setBackgroundColor(HSSFWorkbook workbook, Cell cell, String color) {
        if (Objects.isNull(color)) {
            return;
        }
        HSSFColor.HSSFColorPredefined colorPredefined = colorPredefinedMap.get(color);
        if (Objects.nonNull(colorPredefined)) {
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(colorPredefined.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);
            return;
        }
        if (color.startsWith(HASH)) {
            int r = Integer.parseInt((color.substring(1, 3)), 16);   //转为16进制
            int g = Integer.parseInt((color.substring(3, 5)), 16);
            int b = Integer.parseInt((color.substring(5, 7)), 16);
            //自定义cell颜色
            HSSFPalette palette = workbook.getCustomPalette();
            //这里的9是索引
            palette.setColorAtIndex((short) 999, (byte) r, (byte) g, (byte) b);
            CellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor((short) 999);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(style);
            return;
        }
        if (color.startsWith(RGB)) {
            String rgbColor = color.replace("rgb", "").replace("(", "").replace(")", "");
            String[] rgbColorArr = rgbColor.split(",");
            List<Integer> rgb = Arrays.stream(rgbColorArr)
                    .map(String::trim).map(Integer::parseInt)
                    .collect(Collectors.toList());
            if (rgb.size() != 3) {
                return;
            }
            //自定义cell颜色
            HSSFPalette palette = workbook.getCustomPalette();
            palette.setColorAtIndex((short) 999, rgb.get(0).byteValue(), rgb.get(1).byteValue(), rgb.get(2).byteValue());
            CellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor((short) 999);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(style);
        }
    }
}
