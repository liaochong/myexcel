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

import com.github.liaochong.example.pojo.ArtCrowd;
import com.github.liaochong.myexcel.core.DefaultExcelReader;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.net.URL;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
@RestController
public class DefaultExcelReaderExampleController {

    @GetMapping("/default/excel/read/example")
    public List<ArtCrowd> read() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/templates/read_example.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<ArtCrowd> result = DefaultExcelReader.of(ArtCrowd.class).sheet(0).rowFilter(row -> row.getRowNum() > 0).read(path.toFile());
        return result;
    }
}
