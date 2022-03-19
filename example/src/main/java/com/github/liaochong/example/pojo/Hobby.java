package com.github.liaochong.example.pojo;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;

public class Hobby extends People {
    @ExcelColumn(order = 6, defaultValue = "---")
    String hobby;

    public String getHobby() {
        return hobby;
    }

    public void setHobby(String hobby) {
        this.hobby = hobby;
    }
}