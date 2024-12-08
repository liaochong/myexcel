package com.github.liaochong.myexcel.core.style;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Map;

/**
 * Lock样式
 *
 * @author gzcltech
 * @version 4.5.6
 */
public class LockStyle {
    public static final String LOCK = "lock";

    public static void setLock(CellStyle cellStyle, Map<String, String> tdStyle) {
        String lock = tdStyle.get(LOCK);
        if (Boolean.TRUE.toString().equalsIgnoreCase(lock)) {
            cellStyle.setLocked(true);
        } else if (Boolean.FALSE.toString().equalsIgnoreCase(lock)) {
            cellStyle.setLocked(false);
        }
    }
}
