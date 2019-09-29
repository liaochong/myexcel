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
package com.github.liaochong.myexcel.core.pojo;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.annotation.ExcelTable;
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

import java.math.BigDecimal;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
@ExcelTable(sheetName = "人员信息")
public class WidthPeople {
    @ExcelColumn(title = "姓名", index = 0, style = {"even->color:red;font-weight:bold;width:10px", "odd->color:yellow",
            "title->color:blue;font-weight:bold;font-size:16"})
    String name;

    @ExcelColumn(title = "年龄", index = 1, style = {"even->color:yellow;width:15px;font-weight:bold"})
    Integer age;

    @ExcelColumn(title = "是否会跳舞", index = 2, style = {"even->color:red;width:20px;font-weight:bold"})
    boolean dance;

    @ExcelColumn(title = "金钱", decimalFormat = "#,000.00", index = 3)
    BigDecimal money;
}
