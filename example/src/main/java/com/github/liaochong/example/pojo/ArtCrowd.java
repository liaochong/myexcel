package com.github.liaochong.example.pojo;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.annotation.ExcelTable;

import java.time.LocalDateTime;

/**
 * @author liaochong
 * @version 1.0
 */
@ExcelTable(sheetName = "艺术生", rowAccessWindowSize = 100, useFieldNameAsTitle = true)
public class ArtCrowd extends People {

    @ExcelColumn(order = 3, index = 3, width = 20)
    private String paintingLevel;

    @ExcelColumn(order = 4, title = "是否会跳舞", width = 9, groups = {People.class, String.class}, index = 4)
    private boolean dance;

    @ExcelColumn(order = 5, title = "考核时间", width = 10, dateFormatPattern = "yyyy-MM-dd HH:mm:ss", groups = {People.class, String.class}, index = 5)
    private LocalDateTime assessmentTime;

    @ExcelColumn(order = 6, defaultValue = "---")
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
