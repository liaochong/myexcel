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

import com.github.liaochong.myexcel.utils.FileExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.HashMap;

/**
 * @author liaochong
 * @version 1.0
 */
public class ThymeleafExcelBuilderTest extends BasicTest {

    @Test
    public void build() throws Exception {
        ExcelBuilder excelBuilder = new ThymeleafExcelBuilder();
        Workbook workbook = excelBuilder.classpathTemplate("/templates/thymeleafToExcelExample.html").build(new HashMap<>());
        FileExportUtil.export(workbook, new File(TEST_DIR + "thymeleaf_build.xlsx"));
    }

    @Test
    public void fileBuild() throws Exception {
        ExcelBuilder excelBuilder = new ThymeleafExcelBuilder();
        Workbook workbook = excelBuilder
                .fileTemplate("/Users/liaochong/Develop/Intellij Idea/Workspace/Git/myexcel/src/test/resources/templates", "thymeleafToExcelExample.html")
                .build(new HashMap<>());
        FileExportUtil.export(workbook, new File(TEST_DIR + "thymeleaf_file_build.xlsx"));
    }
}
