/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.data;

import com.miraisolutions.xlconnect.Common;
import com.miraisolutions.xlconnect.utils.CellUtils;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 *
 * @author Martin Studer, Mirai Solutions GmbH
 */
public class ColumnBuilder extends Common {
    // Collection to hold detected data types for each value in a column
    // --> will be used to determine actual final data type for column
    private ArrayList<DataType> detectedTypes;
    // Collection to hold cell references
    private ArrayList<Cell> cells;
    // Collection to hold actual values
    private ArrayList<CellValue> values;

    // TODO: check this
    DateFormat dateFormat = SimpleDateFormat.getDateTimeInstance();

    // Helper collection to store CellValue's that are dates
    // This is needed as a CellValue doesn't store the information whether it is
    // a date or not - dates are just numerics
    private ArrayList<CellValue> isDate = new ArrayList<CellValue>();

    private boolean forceConversion;

    public ColumnBuilder(int nrows, boolean forceConversion) {
        this.detectedTypes = new ArrayList<DataType>(nrows);
        this.cells = new ArrayList<Cell>(nrows);
        this.values = new ArrayList<CellValue>(nrows);
        this.forceConversion = forceConversion;
    }

    public void addMissing() {
        // Add "missing" to collection
        values.add(null);
        // assume "smallest" data type
        detectedTypes.add(DataType.Boolean);
    }

    public void addValue(Cell c, CellValue cv, DataType dt) {
        if(DataType.DateTime.equals(dt)) isDate.add(cv);
        cells.add(c);
        values.add(cv);
        detectedTypes.add(dt);
    }

    public DataType determineColumnType() {
        DataType columnType = DataType.Boolean;

        // Iterate over cell types; as soon as String is detecte we can stop
        Iterator<DataType> it = detectedTypes.iterator();
        while(it.hasNext() && !columnType.equals(DataType.String)) {
            DataType dt = it.next();
            // In case current data type ordinal is bigger than column data type ordinal
            // then adapt column data type to be current data type;
            // this assumes DataType enum to in order from "smallest" to "biggest" data type
            if(dt.ordinal() > columnType.ordinal()) columnType = dt;
        }

        return columnType;
    }

    public ArrayList build(DataType asType) {
        DataType columnType = (asType == null) ? this.determineColumnType() : asType;
        ArrayList colValues = new ArrayList(values.size());
        Iterator<CellValue> it = values.iterator();
        int counter = 0;
        while(it.hasNext()) {
           CellValue cv = it.next();
           if(cv == null) {
               colValues.add(null);
           } /* else if(!CellUtils.isCellValueOfType(cv, columnType)) {
               addWarning("Cell " + CellUtils.formatAsString(cells.get(counter)) + " is not of type '" +
                       columnType.toString() + "' - returning NA");
               colValues.add(null);
           }*/ else {
               switch(columnType) {
                   case Boolean:
                       switch(detectedTypes.get(counter)) {
                           case Boolean:
                               colValues.add(cv.getBooleanValue());
                               break;
                           case Numeric:
                               if(forceConversion)
                                   colValues.add(cv.getNumberValue() > 0);
                               else
                                   colValues.add(null);
                               break;
                           case String:
                               if(forceConversion)
                                   colValues.add(Boolean.valueOf(cv.getStringValue().toLowerCase()));
                               else
                                   colValues.add(null);
                               break;
                           case DateTime:
                               colValues.add(null);
                               addWarning("Cell " + CellUtils.formatAsString(cells.get(counter)) +
                                       " cannot be converted from DateTime to Boolean - returning NA");
                               break;
                           default:
                               throw new IllegalArgumentException("Unknown data type detected!");
                       }
                       break;
                   case DateTime:
                       switch(detectedTypes.get(counter)) {
                           case Boolean:
                               colValues.add(null);
                               addWarning("Cell " + CellUtils.formatAsString(cells.get(counter)) +
                                       " cannot be converted from Boolean to DateTime - returning NA");
                               break;
                           case Numeric:
                               if(forceConversion) {
                                   if(DateUtil.isValidExcelDate(cv.getNumberValue()))
                                       colValues.add(DateUtil.getJavaDate(cv.getNumberValue()));
                                   else {
                                       colValues.add(null);
                                       addWarning("Cell " + CellUtils.formatAsString(cells.get(counter)) +
                                           " cannot be converted from Numeric to DateTime - returning NA");
                                   }
                               } else
                                   colValues.add(null);
                               break;
                           case String:
                               if(forceConversion) {
                                   try {
                                       colValues.add(dateFormat.parse(cv.getStringValue()));
                                   } catch(ParseException e) {
                                       colValues.add(null);
                                       addWarning("Cell " + CellUtils.formatAsString(cells.get(counter)) +
                                           " cannot be converted from String to DateTime - returning NA");
                                   }
                               } else
                                   colValues.add(null);
                               break;
                           case DateTime:
                               colValues.add(DateUtil.getJavaDate(cv.getNumberValue()));
                               break;
                           default:
                               throw new IllegalArgumentException("Unknown data type detected!");
                       }
                       break;
                   case Numeric:
                       switch(detectedTypes.get(counter)) {
                           case Boolean:
                               colValues.add(cv.getBooleanValue() ? 1. : 0.);
                               break;
                           case Numeric:
                               colValues.add(cv.getNumberValue());
                               break;
                           case String:
                               if(forceConversion) {
                                   try {
                                       colValues.add(Double.parseDouble(cv.getStringValue()));
                                   } catch(NumberFormatException e) {
                                       colValues.add(null);
                                       addWarning("Cell " + CellUtils.formatAsString(cells.get(counter)) +
                                           " cannot be converted from String to Numeric - returning NA");
                                   }
                               } else
                                   colValues.add(null);
                               break;
                           case DateTime:
                               if(forceConversion)
                                   colValues.add(cv.getNumberValue());
                               else
                                   colValues.add(null);
                               break;
                           default:
                               throw new IllegalArgumentException("Unknown data type detected!");
                       }
                       break;
                   case String:
                        switch(detectedTypes.get(counter)) {
                           case Boolean:
                               colValues.add(Boolean.toString(cv.getBooleanValue()));
                               break;
                           case Numeric:
                               colValues.add(Double.toString(cv.getNumberValue()));
                               break;
                           case String:
                               colValues.add(cv.getStringValue());
                               break;
                           case DateTime:
                               colValues.add(dateFormat.format(DateUtil.getJavaDate(cv.getNumberValue())));
                               break;
                           default:
                               throw new IllegalArgumentException("Unknown data type detected!");
                       }
                       break;
                   default:
                        throw new IllegalArgumentException("Unknown column type detected!");
               }
           }
           ++counter;
        }

        return colValues;
    }
}
