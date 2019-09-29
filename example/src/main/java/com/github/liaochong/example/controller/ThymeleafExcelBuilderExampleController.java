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
package com.github.liaochong.example.controller;

import com.github.liaochong.myexcel.core.ExcelBuilder;
import com.github.liaochong.myexcel.core.ThymeleafExcelBuilder;
import com.github.liaochong.myexcel.utils.AttachmentExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author liaochong
 * @version 1.0
 */
@RestController
public class ThymeleafExcelBuilderExampleController {

    @GetMapping("/thymeleaf/example")
    public void build(HttpServletResponse response) throws IOException {
        try (ExcelBuilder excelBuilder = new ThymeleafExcelBuilder()) {
            Workbook workbook = excelBuilder.template("/templates/demo2.html").build(new HashMap<>());
            AttachmentExportUtil.export(workbook, "thymeleaf_excel", response);
        }
    }
}
