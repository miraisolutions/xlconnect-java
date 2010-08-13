/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.integration.r;

import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import java.util.Arrays;
import java.util.Date;
import java.util.Vector;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public final class RDataFrameWrapper {

    final DataFrame dataFrame;

    public RDataFrameWrapper() {
        this.dataFrame = new DataFrame();
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

    public void addDateTimeColumn(String name, Date[] column, boolean[] na) {
        // TODO: still to be implemented
    }
}
