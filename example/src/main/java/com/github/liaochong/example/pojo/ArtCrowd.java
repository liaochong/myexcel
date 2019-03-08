package com.github.liaochong.example.pojo;

import com.github.liaochong.myexcel.core.WorkbookType;
import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.annotation.ExcelTable;
import com.github.liaochong.myexcel.core.annotation.ExcludeColumn;

import java.time.LocalDateTime;

/**
 * @author liaochong
 * @version 1.0
 */
@ExcelTable(sheetName = "艺术生", workbookType = WorkbookType.SXLSX, rowAccessWindowSize = 100, useFieldNameAsTitle = true)
public class ArtCrowd extends People {

    @ExcelColumn(order = 4)
    private String paintingLevel;

    @ExcelColumn(order = 5, title = "是否会跳舞", groups = {People.class, String.class})
    private boolean dance;

    @ExcelColumn(order = 6, title = "考核时间", dateFormatPattern = "yyyy-MM-dd HH:mm:ss", groups = {People.class, String.class})
    private LocalDateTime assessmentTime;

    @ExcludeColumn
    private String hobby;

    public String getPaintingLevel() {
        return paintingLevel;
    }

    public void setPaintingLevel(String paintingLevel) {
        this.paintingLevel = paintingLevel;
    }

    public boolean isDance() {
        return dance;
    }

    public void setDance(boolean dance) {
        this.dance = dance;
    }

    public LocalDateTime getAssessmentTime() {
        return assessmentTime;
    }

    public void setAssessmentTime(LocalDateTime assessmentTime) {
        this.assessmentTime = assessmentTime;
    }

    public String getHobby() {
        return hobby;
    }

    public void setHobby(String hobby) {
        this.hobby = hobby;
    }
}
