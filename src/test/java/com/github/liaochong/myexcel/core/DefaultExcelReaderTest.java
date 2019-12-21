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

import com.github.liaochong.myexcel.core.pojo.Picture;
import org.junit.jupiter.api.Test;

import java.net.URL;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
public class DefaultExcelReaderTest extends BasicTest {

    @Test
    public void readXlsxPicture() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/picture.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());
        List<Picture> pictures = DefaultExcelReader.of(Picture.class).read(path.toFile());
        System.out.println("");
    }

    @Test
    public void readXlsPicture() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/picture.xls");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());
        List<Picture> pictures = DefaultExcelReader.of(Picture.class).read(path.toFile());
        System.out.println("");
    }
}
