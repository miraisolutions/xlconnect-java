/*
 *
    XLConnect
    Copyright (C) 2010-2024 Mirai Solutions GmbH

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

package com.miraisolutions.xlconnect.integration.r;

import com.miraisolutions.xlconnect.data.Column;
import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;

import java.util.Arrays;
import java.util.BitSet;
import java.util.Date;
import java.util.stream.IntStream;

public final class RDataFrameWrapper {

    final DataFrame dataFrame;

    public RDataFrameWrapper() {
        this.dataFrame = new DataFrame();
    }

    public RDataFrameWrapper(DataFrame dataFrame) {
        this.dataFrame = dataFrame;
    }

    public void addNumericColumn(String name, double[] column, boolean[] na) {
        dataFrame.addColumn(name, new Column(column, toBitSet(na), DataType.Numeric));
    }

    public void addBooleanColumn(String name, boolean[] column, boolean[] na) {
        dataFrame.addColumn(name, new Column(column, toBitSet(na), DataType.Boolean));
    }

    public void addStringColumn(String name, String[] column, boolean[] na) {
        dataFrame.addColumn(name, new Column(column, toBitSet(na), DataType.String));
    }

    public void addDateTimeColumn(String name, long[] column, boolean[] na) {
        Date[] elements = IntStream.range(0, column.length)
                .mapToObj(i -> na[i] ? null : new Date(column[i]))
                .toArray(Date[]::new);
        dataFrame.addColumn(name, new Column(elements, toBitSet(na), DataType.DateTime));
    }

    public String[] getColumnTypes() {
        return dataFrame.getColumnTypes().stream()
                .map(DataType::toString)
                .toArray(String[]::new);
    }

    public String[] getColumnNames() {
        return dataFrame.getColumnNames().toArray(String[]::new);
    }

    public double[] getNumericColumn(int col) {
        return dataFrame.getColumn(col).getNumericData();
    }

    public String[] getStringColumn(int col) {
        return dataFrame.getColumn(col).getStringData();
    }

    public boolean[] getBooleanColumn(int col) {
        return dataFrame.getColumn(col).getBooleanData();
    }

    public long[] getDateTimeColumn(int col) {
        return Arrays.stream(dataFrame.getColumn(col).getDateTimeData())
                .mapToLong(date -> date == null ? 0 : date.getTime())
                .toArray();
    }

    public boolean[] isMissing(int col) {
        BitSet missing = dataFrame.getColumn(col).getMissing();
        boolean[] na = new boolean[missing.length()];
        missing.stream().forEach(i -> na[i] = true);
        return na;
    }

    private static BitSet toBitSet(boolean[] bits) {
        BitSet bs = new BitSet(bits.length);
        for (int i = 0; i < bits.length; i++) {
            bs.set(i, bits[i]);
        }
        return bs;
    }
}
