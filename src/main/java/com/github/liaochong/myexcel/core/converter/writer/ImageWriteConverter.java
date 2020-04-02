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
package com.github.liaochong.myexcel.core.converter.writer;

import com.github.liaochong.myexcel.core.ConvertContext;
import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.constant.FileType;
import com.github.liaochong.myexcel.core.constant.ImageFile;
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.converter.WriteConverter;

import java.io.File;
import java.lang.reflect.Field;

/**
 * 图片写转换器
 *
 * @author liaochong
 * @version 1.0
 */
public class ImageWriteConverter implements WriteConverter {

    @Override
    public Pair<Class, Object> convert(Field field, Object fieldVal, ConvertContext convertContext) {
        return Pair.of(ImageFile.class, fieldVal);
    }

    @Override
    public boolean support(Field field, Object fieldVal, ConvertContext convertContext) {
        if (field.getType() != File.class) {
            return false;
        }
        ExcelColumnMapping mapping = convertContext.getExcelColumnMappingMap().get(field);
        return mapping != null && mapping.getFileType() == FileType.IMAGE;
    }
}
