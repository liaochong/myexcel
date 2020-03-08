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

import com.github.liaochong.myexcel.core.pojo.Product;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
public class BeetlExcelBuilderTest extends BasicTest {

    @Test
    public void build() throws Exception {
        ExcelBuilder excelBuilder = new BeetlExcelBuilder();
        Workbook workbook = excelBuilder.classpathTemplate("/templates/beetlToExcelExample.btl").build(getDataMap());
        FileExportUtil.export(workbook, new File(TEST_DIR + "beetl_build.xlsx"));
    }

    @Test
    public void fileBuild() throws Exception {
        ExcelBuilder excelBuilder = new BeetlExcelBuilder();
        Workbook workbook = excelBuilder
                .fileTemplate("/Users/liaochong/Develop/Intellij Idea/Workspace/Git/myexcel/src/test/resources/templates", "beetlToExcelExample.btl")
                .build(getDataMap());
        FileExportUtil.export(workbook, new File(TEST_DIR + "beetl_file_build.xlsx"));
    }

    private Map<String, Object> getDataMap() {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("sheetName", "beetl_excel_example");

        List<String> titles = new ArrayList<>();
        titles.add("Category");
        titles.add("Product Name");
        titles.add("Count");
        dataMap.put("titles", titles);

        List<Product> data = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Product product = new Product();
            if (i % 2 == 0) {
                product.setCategory("蔬菜");
                product.setName("小白菜");
                product.setCount(100);
            } else {
                product.setCategory("电子产品");
                product.setName("ipad");
                product.setCount(999);
            }
            data.add(product);
        }
        dataMap.put("data", data);
        return dataMap;
    }
}
