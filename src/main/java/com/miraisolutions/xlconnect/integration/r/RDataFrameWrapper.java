/*
 *
    XLConnect
    Copyright (C) 2010 Mirai Solutions GmbH

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
        dataFrame.addColumn(name, new Column(column, na, DataType.Numeric));
    }

    public void addBooleanColumn(String name, boolean[] column, boolean[] na) {
        dataFrame.addColumn(name, new Column(column, na, DataType.Boolean));
    }

    public void addStringColumn(String name, String[] column, boolean[] na) {
        dataFrame.addColumn(name, new Column(column, na, DataType.String));
    }

    public void addDateTimeColumn(String name, long[] column, boolean[] na) {
        Date[] elements = new Date[column.length];
        for(int i = 0; i < column.length; i++) {
            if(na[i])
                elements[i] = null;
            else
                elements[i] = new Date(column[i]);
        }
        dataFrame.addColumn(name, new Column(elements, na, DataType.DateTime));
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
        return columnNames.toArray(new String[columnNames.size()]);
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
            if(v[i] == null)
                values[i] = 0;
            else
                values[i] = v[i].getTime();
        }

        return values;
    }

    public boolean[] isMissing(int col) {
        return dataFrame.getColumn(col).getMissing();
    }
}
