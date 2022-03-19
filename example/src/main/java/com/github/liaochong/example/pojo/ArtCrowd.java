package com.github.liaochong.example.pojo;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.annotation.ExcelModel;

import java.time.LocalDateTime;

/**
 * @author liaochong
 * @version 1.0
 */
@ExcelModel(sheetName = "data", useFieldNameAsTitle = true)
public class ArtCrowd extends People {

    private final Hobby hobby = new Hobby();
    @ExcelColumn(order = 3, index = 3)
    private String paintingLevel;

    @ExcelColumn(order = 4, title = "Dance", groups = {People.class, String.class}, index = 4)
    private boolean dance;

    @ExcelColumn(order = 5, title = "Time", groups = {People.class, String.class}, index = 5)
    private LocalDateTime assessmentTime;

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

}
