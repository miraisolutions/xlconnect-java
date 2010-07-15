/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.data;

import java.util.Vector;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class DataFrame {
    
    protected Vector<String> columnNames;
    protected Vector<DataType> columnTypes;
    protected Vector<Vector> columns;

    public DataFrame() {
        this.columnNames = new Vector<String>();
        this.columnTypes = new Vector<DataType>();
        this.columns = new Vector<Vector>();
    }

    public int columns() {
        return columns.size();
    }

    public int rows() {
        if(isEmpty())
            return 0;
        else
            return columns.get(0).size();
    }

    public boolean isEmpty() {
        return columns.isEmpty();
    }

    public boolean hasColumnHeader() {
        boolean hasHeader = false;
        for(int i = 0; i < columnNames.size(); i++) {
            if(columnNames.get(i) != null) {
                hasHeader = true;
                break;
            }
        }

        return hasHeader;
    }

    
    public void addColumn(String name, DataType type, Vector column) {
        if(isEmpty() || (column.size() == rows())) {
            columnNames.add(name);
            columnTypes.add(type);
            columns.add(column);
        } else
            throw new IllegalArgumentException("Length of specified column does not match length of " +
                    "existing columns in the DataFrame!");
    }

    public String getColumnName(int index) {
        return columnNames.get(index);
    }

    public DataType getColumnType(int index) {
        return columnTypes.get(index);
    }

    public Vector getColumn(int index) {
        return columns.get(index);
    }

    /*
    public Vector getColumn(String name) {
        int index = columnNames.indexOf(name);
        if(index >= 0)
            return getColumn(index);
        else
            throw new IllegalArgumentException("No column '" + name + "' available!");

    }
     *
     */

}
