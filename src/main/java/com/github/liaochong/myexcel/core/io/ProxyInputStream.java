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

import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author liaochong
 * @version 1.0
 */
public abstract class ProxyInputStream extends FilterInputStream {
    public ProxyInputStream(InputStream proxy) {
        super(proxy);
    }

    @Override
    public int read() throws IOException {
        try {
            this.beforeRead(1);
            int b = this.in.read();
            this.afterRead(b != -1 ? 1 : -1);
            return b;
        } catch (IOException e) {
            this.handleIOException(e);
            return -1;
        }
    }

    @Override
    public int read(byte[] bts) throws IOException {
        try {
            this.beforeRead(bts == null ? 0 : bts.length);
            int n = this.in.read(bts);
            this.afterRead(n);
            return n;
        } catch (IOException e) {
            this.handleIOException(e);
            return -1;
        }
    }

    @Override
    public int read(byte[] bts, int off, int len) throws IOException {
        try {
            this.beforeRead(len);
            int n = this.in.read(bts, off, len);
            this.afterRead(n);
            return n;
        } catch (IOException e) {
            this.handleIOException(e);
            return -1;
        }
    }

    @Override
    public long skip(long ln) throws IOException {
        try {
            return this.in.skip(ln);
        } catch (IOException e) {
            this.handleIOException(e);
            return 0L;
        }
    }

    @Override
    public int available() throws IOException {
        try {
            return super.available();
        } catch (IOException e) {
            this.handleIOException(e);
            return 0;
        }
    }

    @Override
    public void close() throws IOException {
        if (this.in != null) {
            try {
                this.in.close();
            } catch (IOException e) {
                this.handleIOException(e);
            }
        }
    }

    @Override
    public synchronized void mark(int readlimit) {
        this.in.mark(readlimit);
    }

    @Override
    public synchronized void reset() throws IOException {
        try {
            this.in.reset();
        } catch (IOException e) {
            this.handleIOException(e);
        }

    }

    @Override
    public boolean markSupported() {
        return this.in.markSupported();
    }

    protected void beforeRead(int n) throws IOException {
    }

    protected void afterRead(int n) throws IOException {
    }

    protected void handleIOException(IOException e) throws IOException {
        throw e;
    }
}
