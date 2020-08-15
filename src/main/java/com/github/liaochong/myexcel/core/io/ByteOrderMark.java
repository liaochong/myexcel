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
package com.github.liaochong.myexcel.core.io;

import java.util.Locale;

/**
 * @author liaochong
 * @version 1.0
 */
public class ByteOrderMark {
    private static final long serialVersionUID = 1L;
    public static final ByteOrderMark UTF_8 = new ByteOrderMark("UTF-8", new int[]{239, 187, 191});
    public static final ByteOrderMark UTF_16BE = new ByteOrderMark("UTF-16BE", new int[]{254, 255});
    public static final ByteOrderMark UTF_16LE = new ByteOrderMark("UTF-16LE", new int[]{255, 254});
    public static final ByteOrderMark UTF_32BE = new ByteOrderMark("UTF-32BE", new int[]{0, 0, 254, 255});
    public static final ByteOrderMark UTF_32LE = new ByteOrderMark("UTF-32LE", new int[]{255, 254, 0, 0});
    public static final char UTF_BOM = '\ufeff';
    private final String charsetName;
    private final int[] bytes;

    public ByteOrderMark(String charsetName, int... bytes) {
        if (charsetName != null && !charsetName.isEmpty()) {
            if (bytes != null && bytes.length != 0) {
                this.charsetName = charsetName;
                this.bytes = new int[bytes.length];
                System.arraycopy(bytes, 0, this.bytes, 0, bytes.length);
            } else {
                throw new IllegalArgumentException("No bytes specified");
            }
        } else {
            throw new IllegalArgumentException("No charsetName specified");
        }
    }

    public String getCharsetName() {
        return this.charsetName;
    }

    public int length() {
        return this.bytes.length;
    }

    public int get(int pos) {
        return this.bytes[pos];
    }

    public byte[] getBytes() {
        byte[] copy = new byte[this.bytes.length];

        for (int i = 0; i < this.bytes.length; ++i) {
            copy[i] = (byte) this.bytes[i];
        }

        return copy;
    }

    @Override
    public boolean equals(Object obj) {
        if (!(obj instanceof ByteOrderMark)) {
            return false;
        } else {
            ByteOrderMark bom = (ByteOrderMark) obj;
            if (this.bytes.length != bom.length()) {
                return false;
            } else {
                for (int i = 0; i < this.bytes.length; ++i) {
                    if (this.bytes[i] != bom.get(i)) {
                        return false;
                    }
                }

                return true;
            }
        }
    }

    @Override
    public int hashCode() {
        int hashCode = this.getClass().hashCode();
        int[] var2 = this.bytes;
        int var3 = var2.length;

        for (int var4 = 0; var4 < var3; ++var4) {
            int b = var2[var4];
            hashCode += b;
        }

        return hashCode;
    }

    @Override
    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append(this.getClass().getSimpleName());
        builder.append('[');
        builder.append(this.charsetName);
        builder.append(": ");

        for (int i = 0; i < this.bytes.length; ++i) {
            if (i > 0) {
                builder.append(",");
            }

            builder.append("0x");
            builder.append(Integer.toHexString(255 & this.bytes[i]).toUpperCase(Locale.ROOT));
        }

        builder.append(']');
        return builder.toString();
    }
}
