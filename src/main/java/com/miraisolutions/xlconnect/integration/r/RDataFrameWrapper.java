/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.integration.r;

import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Vector;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public final class RDataFrameWrapper {

    final DataFrame dataFrame;
    private final static SimpleDateFormat dateTimeParser = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

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
        Vector<Double> v = new Vector<Double>(Arrays.asList(elements));
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
        Vector<Boolean> v = new Vector<Boolean>(Arrays.asList(elements));
        dataFrame.addColumn(name, DataType.Boolean, v);
    }

    public void addStringColumn(String name, String[] column, boolean[] na) {
        for(int i = 0; i < column.length; i++) {
            if(na[i]) column[i] = null;
        }
        Vector<String> v = new Vector<String>(Arrays.asList(column));
        dataFrame.addColumn(name, DataType.String, v);
    }

    public void addDateTimeColumn(String name, String[] column, boolean[] na) throws ParseException {
        Date[] elements = new Date[column.length];
        for(int i = 0; i < column.length; i++) {
            if(na[i])
                elements[i] = null;
            else 
                elements[i] = dateTimeParser.parse(column[i]);
        }
    }

    public String[] getColumnTypes() {
        Vector<DataType> columnTypes = dataFrame.getColumnTypes();
        String[] dataTypes = new String[columnTypes.size()];
        for(int i = 0; i < columnTypes.size(); i++) {
            dataTypes[i] = columnTypes.get(i).toString();
        }
        return dataTypes;
    }

    public String[] getColumnNames() {
        Vector<String> columnNames = dataFrame.getColumnNames();
        return columnNames.toArray(new String[columnNames.size()]);
    }

    public double[] getNumericColumn(int col) {
        Vector<Double> v = dataFrame.getColumn(col);
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
        Vector<String> v = dataFrame.getColumn(col);
        return v.toArray(new String[v.size()]);
    }

    public boolean[] getBooleanColumn(int col) {
        Vector<Boolean> v = dataFrame.getColumn(col);
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
        Vector<Date> v = dataFrame.getColumn(col);
        String[] values = new String[v.size()];

        for(int i = 0; i < v.size(); i++) {
            Date d = v.get(i);
            if(d == null)
                values[i] = null;
            else
                values[i] = dateTimeParser.format(d);
        }

        return values;
    }

    public boolean[] isMissing(int col) {
        Vector v = dataFrame.getColumn(col);
        boolean[] missing = new boolean[v.size()];
        for(int i = 0; i < v.size(); i++) {
            missing[i] = v.get(i) == null;
        }
        return missing;
    }
}
