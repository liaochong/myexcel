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
package com.github.liaochong.myexcel.utils;

import com.github.liaochong.myexcel.core.constant.Constants;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Base64;

/**
 * @author liaochong
 * @version 1.0
 */
public final class ImageUtil {

    /**
     * 根据地址获得数据的字节流
     *
     * @param strUrl 网络连接地址
     * @return 图片流
     */
    public static InputStream getImageFromNetByUrl(String strUrl) {
        if (StringUtil.isBlank(strUrl)) {
            return null;
        }
        try {
            URL url = new URL(strUrl);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(5 * 1000);
            return conn.getInputStream();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static InputStream generateImageFromBase64(String imgData) {
        try {
            // Base64解码
            byte[] b = Base64.getDecoder().decode(imgData.substring(imgData.indexOf(Constants.COMMA) + 1));
            return new ByteArrayInputStream(b);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
