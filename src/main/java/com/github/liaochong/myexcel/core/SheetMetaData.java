/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core;

/**
 * sheet元数据
 *
 * @author liaochong
 * @version 1.0
 */
public class SheetMetaData {

    private String sheetName;

    private int sheetIndex = -1;

    private int lastRowNum = -1;

    public SheetMetaData(String sheetName, int sheetIndex) {
        this.sheetName = sheetName;
        this.sheetIndex = sheetIndex;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getLastRowNum() {
        return lastRowNum;
    }

    public void setLastRowNum(int lastRowNum) {
        this.lastRowNum = lastRowNum;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }
}
