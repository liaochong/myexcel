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

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
public class BOMInputStream extends ProxyInputStream {
    private final boolean include;
    private final List<ByteOrderMark> boms;
    private ByteOrderMark byteOrderMark;
    private int[] firstBytes;
    private int fbLength;
    private int fbIndex;
    private int markFbIndex;
    private boolean markedAtStart;
    private static final Comparator<ByteOrderMark> ByteOrderMarkLengthComparator = (bom1, bom2) -> {
        int len1 = bom1.length();
        int len2 = bom2.length();
        if (len1 > len2) {
            return -1;
        } else {
            return len2 > len1 ? 1 : 0;
        }
    };

    public BOMInputStream(InputStream delegate) {
        this(delegate, false, ByteOrderMark.UTF_8);
    }

    public BOMInputStream(InputStream delegate, boolean include) {
        this(delegate, include, ByteOrderMark.UTF_8);
    }

    public BOMInputStream(InputStream delegate, ByteOrderMark... boms) {
        this(delegate, false, boms);
    }

    public BOMInputStream(InputStream delegate, boolean include, ByteOrderMark... boms) {
        super(delegate);
        if ((boms == null ? 0 : boms.length) == 0) {
            throw new IllegalArgumentException("No BOMs specified");
        } else {
            this.include = include;
            List<ByteOrderMark> list = Arrays.asList(boms);
            Collections.sort(list, ByteOrderMarkLengthComparator);
            this.boms = list;
        }
    }

    public boolean hasBOM() throws IOException {
        return this.getBOM() != null;
    }

    public boolean hasBOM(ByteOrderMark bom) throws IOException {
        if (!this.boms.contains(bom)) {
            throw new IllegalArgumentException("Stream not configure to detect " + bom);
        } else {
            this.getBOM();
            return this.byteOrderMark != null && this.byteOrderMark.equals(bom);
        }
    }

    public ByteOrderMark getBOM() throws IOException {
        if (this.firstBytes == null) {
            this.fbLength = 0;
            int maxBomSize = ((ByteOrderMark) this.boms.get(0)).length();
            this.firstBytes = new int[maxBomSize];

            for (int i = 0; i < this.firstBytes.length; ++i) {
                this.firstBytes[i] = this.in.read();
                ++this.fbLength;
                if (this.firstBytes[i] < 0) {
                    break;
                }
            }
            this.byteOrderMark = this.find();
            if (this.byteOrderMark != null && !this.include) {
                if (this.byteOrderMark.length() < this.firstBytes.length) {
                    this.fbIndex = this.byteOrderMark.length();
                } else {
                    this.fbLength = 0;
                }
            }
        }
        return this.byteOrderMark;
    }

    public String getBOMCharsetName() throws IOException {
        this.getBOM();
        return this.byteOrderMark == null ? null : this.byteOrderMark.getCharsetName();
    }

    private int readFirstBytes() throws IOException {
        this.getBOM();
        return this.fbIndex < this.fbLength ? this.firstBytes[this.fbIndex++] : -1;
    }

    private ByteOrderMark find() {
        Iterator iterator = this.boms.iterator();

        ByteOrderMark bom;
        do {
            if (!iterator.hasNext()) {
                return null;
            }
            bom = (ByteOrderMark) iterator.next();
        } while (!this.matches(bom));

        return bom;
    }

    private boolean matches(ByteOrderMark bom) {
        for (int i = 0; i < bom.length(); ++i) {
            if (bom.get(i) != this.firstBytes[i]) {
                return false;
            }
        }
        return true;
    }

    @Override
    public int read() throws IOException {
        int b = this.readFirstBytes();
        return b >= 0 ? b : this.in.read();
    }

    @Override
    public int read(byte[] buf, int off, int len) throws IOException {
        int firstCount = 0;
        int b = 0;

        while (len > 0 && b >= 0) {
            b = this.readFirstBytes();
            if (b >= 0) {
                buf[off++] = (byte) (b & 255);
                --len;
                ++firstCount;
            }
        }

        int secondCount = this.in.read(buf, off, len);
        return secondCount < 0 ? (firstCount > 0 ? firstCount : -1) : firstCount + secondCount;
    }

    @Override
    public int read(byte[] buf) throws IOException {
        return this.read(buf, 0, buf.length);
    }

    @Override
    public synchronized void mark(int readlimit) {
        this.markFbIndex = this.fbIndex;
        this.markedAtStart = this.firstBytes == null;
        this.in.mark(readlimit);
    }

    @Override
    public synchronized void reset() throws IOException {
        this.fbIndex = this.markFbIndex;
        if (this.markedAtStart) {
            this.firstBytes = null;
        }
        this.in.reset();
    }

    @Override
    public long skip(long n) throws IOException {
        int skipped;
        for (skipped = 0; n > (long) skipped && this.readFirstBytes() >= 0; ++skipped) {
        }
        return this.in.skip(n - (long) skipped) + (long) skipped;
    }
}
