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

import com.miraisolutions.xlconnect.Workbook;
import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import java.util.Arrays;
import java.util.Date;
import java.util.ArrayList;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public final class RDataFrameWrapper {

    final DataFrame dataFrame;

    public RDataFrameWrapper() {
        this.dataFrame = new DataFrame();
    }

    public RDataFrameWrapper(DataFrame dataFrame) {
        this.dataFrame = dataFrame;
    }

    public void addNumericColumn(String name, double[] column, boolean[] na) {
        Double[] elements = new Double[column.length];
        for(int i = 0; i < column.length; i++) {
            if(na[i])
                elements[i] = null;
            else
                elements[i] = new Double(column[i]);
        }
        ArrayList<Double> v = new ArrayList<Double>(Arrays.asList(elements));
        dataFrame.addColumn(name, DataType.Numeric, v);
    }

    public void addBooleanColumn(String name, boolean[] column, boolean[] na) {
        Boolean[] elements = new Boolean[column.length];
        for(int i = 0; i < column.length; i++) {
            if(na[i])
                elements[i] = null;
            else
                elements[i] = new Boolean(column[i]);
        }
        ArrayList<Boolean> v = new ArrayList<Boolean>(Arrays.asList(elements));
        dataFrame.addColumn(name, DataType.Boolean, v);
    }

    public void addStringColumn(String name, String[] column, boolean[] na) {
        for(int i = 0; i < column.length; i++) {
            if(na[i]) column[i] = null;
        }
        ArrayList<String> v = new ArrayList<String>(Arrays.asList(column));
        dataFrame.addColumn(name, DataType.String, v);
    }

    public void addDateTimeColumn(String name, String[] column, boolean[] na) {
        Date[] elements = new Date[column.length];
        for(int i = 0; i < column.length; i++) {
            if(na[i])
                elements[i] = null;
            else 
                elements[i] = Workbook.dateTimeFormatter.parse(column[i], Workbook.DATE_TIME_FORMAT);
        }
        ArrayList<Date> v = new ArrayList<Date>(Arrays.asList(elements));
        dataFrame.addColumn(name, DataType.DateTime, v);
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
        ArrayList<Double> v = dataFrame.getColumn(col);
        double[] values = new double[v.size()];

        for(int i = 0; i < v.size(); i++) {
            Double d = v.get(i);
            if(d == null)
                values[i] = 0.0;
            else
                values[i] = d.doubleValue();
        }

        return values;
    }

    public String[] getStringColumn(int col) {
        ArrayList<String> v = dataFrame.getColumn(col);
        return v.toArray(new String[v.size()]);
    }

    public boolean[] getBooleanColumn(int col) {
        ArrayList<Boolean> v = dataFrame.getColumn(col);
        boolean[] values = new boolean[v.size()];

        for(int i = 0; i < v.size(); i++) {
            Boolean b = v.get(i);
            if(b == null)
                values[i] = false;
            else
                values[i] = b.booleanValue();
        }

        return values;
    }

    public String[] getDateTimeColumn(int col) {
        ArrayList<Date> v = dataFrame.getColumn(col);
        String[] values = new String[v.size()];

        for(int i = 0; i < v.size(); i++) {
            Date d = v.get(i);
            if(d == null)
                values[i] = null;
            else
                values[i] = Workbook.dateTimeFormatter.format(d, Workbook.DATE_TIME_FORMAT);
        }

        return values;
    }

    public boolean[] isMissing(int col) {
        ArrayList v = dataFrame.getColumn(col);
        boolean[] missing = new boolean[v.size()];
        for(int i = 0; i < v.size(); i++) {
            missing[i] = v.get(i) == null;
        }
        return missing;
    }
}
