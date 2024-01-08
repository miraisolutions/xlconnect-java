/*
 *
    XLConnect
    Copyright (C) 2010-2018 Mirai Solutions GmbH

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
import java.util.ArrayList;
import java.util.BitSet;
import java.util.Date;

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
        Date[] elements = new Date[column.length];
        for (int i = 0; i < column.length; i++) {
            elements[i] = na[i] ? null : new Date(column[i]);
        }
        dataFrame.addColumn(name, new Column(elements, toBitSet(na), DataType.DateTime));
    }

    public String[] getColumnTypes() {
        ArrayList<DataType> columnTypes = dataFrame.getColumnTypes();
        String[] dataTypes = new String[columnTypes.size()];
        for(int i = 0; i < columnTypes.size(); i++) {
            dataTypes[i] = columnTypes.get(i).toString();
        }
        return dataTypes;
    }

    public String[] getColumnNames() {
        ArrayList<String> columnNames = dataFrame.getColumnNames();
        return columnNames.toArray(new String[0]);
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
        Date[] v = dataFrame.getColumn(col).getDateTimeData();
        long[] values = new long[v.length];

        for(int i = 0; i < v.length; i++) {
            values[i] = (v[i] == null) ? 0 : v[i].getTime();
        }

        return values;
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
