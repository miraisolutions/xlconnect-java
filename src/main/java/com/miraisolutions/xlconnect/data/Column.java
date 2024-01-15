/*
 *
    XLConnect
    Copyright (C) 2013-2024 Mirai Solutions GmbH

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *
 */
package com.miraisolutions.xlconnect.data;

import java.util.BitSet;
import java.util.Date;

public final class Column {
    private final Object data;
    private final int size;
    private final BitSet missing;
    private final DataType type;

    public Column(Object data, int size, BitSet missing, DataType type) {
        this.data = data;
        this.size = size;
        this.missing = missing;
        this.type = type;
    }

    public DataType getDataType() {
        return type;
    }

    public boolean[] getBooleanData() {
        return (boolean[]) data;
    }

    public Date[] getDateTimeData() {
        return (Date[]) data;
    }

    public double[] getNumericData() {
        return (double[]) data;
    }

    public String[] getStringData() {
        return (String[]) data;
    }

    public BitSet getMissing() {
        return missing;
    }

    public boolean isMissing(int i) {
        return missing.get(i);
    }

    public int size() {
        return size;
    }
}
