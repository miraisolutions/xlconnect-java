/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect.data;

import com.miraisolutions.xlconnect.Common;
import com.miraisolutions.xlconnect.ErrorBehavior;
import com.miraisolutions.xlconnect.Workbook;
import com.miraisolutions.xlconnect.utils.CellUtils;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

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
    // Date/time format used for conversion to and from strings
    private String dateTimeFormat;

    // Helper collection to store CellValues that are dates
    // This is needed as a CellValue doesn't store the information whether it is
    // a date or not - dates are just numerics
    private ArrayList<CellValue> isDate = new ArrayList<CellValue>();
    // Should conversion to a less generic data type be forced?
    private boolean forceConversion;

    private boolean takeCached = false;
    private FormulaEvaluator evaluator = null;
    private ErrorBehavior onErrorCell;
    String[] missingValue;

    public ColumnBuilder(int nrows, boolean forceConversion,
            FormulaEvaluator evaluator, ErrorBehavior onErrorCell,
            String[] missingValue) {
        this.detectedTypes = new ArrayList<DataType>(nrows);
        this.cells = new ArrayList<Cell>(nrows);
        this.values = new ArrayList<CellValue>(nrows);
        this.forceConversion = forceConversion;
        this.evaluator = evaluator;
        this.takeCached = evaluator == null;
        this.onErrorCell = onErrorCell;
        this.missingValue = missingValue;
    }

    public void addMissing() {
        // Add "missing" to collection
        values.add(null);
        cells.add(null);
        // assume "smallest" data type
        detectedTypes.add(DataType.Boolean);
    }


    private void cellError(String msg) {
        if(this.onErrorCell.equals(ErrorBehavior.WARN)) {
            this.addMissing();
            this.addWarning(msg);
        } else
            throw new IllegalArgumentException(msg);
    }


    // extracts the cached value from a cell without re-evaluating
    // the formula. returns null if the cell is blank.
    private CellValue getCachedCellValue(Cell cell) {
        int valueType = cell.getCellType();
        if(valueType == Cell.CELL_TYPE_FORMULA) {
            valueType = cell.getCachedFormulaResultType();
        }

        switch(valueType) {
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_BOOLEAN:
                if(cell.getBooleanCellValue()) {
                    return CellValue.TRUE;
                } else {
                    return CellValue.FALSE;
                }
            case Cell.CELL_TYPE_NUMERIC:
                return new CellValue(cell.getNumericCellValue());
            case Cell.CELL_TYPE_STRING:
                return new CellValue(cell.getStringCellValue());
            case Cell.CELL_TYPE_ERROR:
                return CellValue.getError(cell.getErrorCellValue());
            default:
                String msg =  String.format("Could not extract value from cell with cached value type %d", valueType);
                throw new RuntimeException(msg);
        }
    }

    // extracts the value from a cell by either evaluating it or taking the
    // cached value
    private CellValue getCellValue(Cell cell) {
        if(this.takeCached) {
            return getCachedCellValue(cell);
        } else {
            return this.evaluator.evaluate(cell);
        }
    }

    public void addCell(Cell c) {
        String msg;

        // In case the cell does not exist ...
        if(c == null) {
            this.addMissing();
            return;
        }

       /*
         * The following is to handle error cells (before they have been evaluated
         * to a CellValue) and cells which are formulas but have cached errors.
         */
        if(
            c.getCellType() == Cell.CELL_TYPE_ERROR ||
            (c.getCellType() == Cell.CELL_TYPE_FORMULA &&
             c.getCachedFormulaResultType() == Cell.CELL_TYPE_ERROR)
        ) {
             msg = "Error detected in cell " + CellUtils.formatAsString(c) +
                    " - " + CellUtils.getErrorMessage(c.getErrorCellValue());
            cellError(msg);
            return;
        }

        CellValue cv = null;

        // Try to evaluate cell;
        // report an error if this fails
        try {
            cv = getCellValue(c);
        } catch(Exception e) {
            msg = "Error when trying to evaluate cell " + 
                    CellUtils.formatAsString(c) + " - " + e.getMessage();
            cellError(msg);
            return;
        }

        // Not sure if this case should ever happen;
        // let's be sure anyway
        if(cv == null){
            addMissing();
            return;
        }

       // Determine (evaluated) cell data type
        switch(cv.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                addMissing();
                return;
            case Cell.CELL_TYPE_BOOLEAN:
                addValue(c, cv, DataType.Boolean);
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if(DateUtil.isCellDateFormatted(c))
                    addValue(c, cv, DataType.DateTime);
                else
                    addValue(c, cv, DataType.Numeric);
                break;
            case Cell.CELL_TYPE_STRING:
                boolean missing = false;
                for(int i = 0; i < missingValue.length; i++) {
                    if(cv.getStringValue() == null || cv.getStringValue().equals(missingValue[i])) {
                        missing = true;
                        break;
                    }
                }
                if(missing)
                    addMissing();
                else
                    addValue(c, cv, DataType.String);
                break;
            case Cell.CELL_TYPE_FORMULA:
                msg = "Formula detected in already evaluated cell " + CellUtils.formatAsString(c) + "!";
                cellError(msg);
                break;
            case Cell.CELL_TYPE_ERROR:
                msg = "Error detected in cell " + CellUtils.formatAsString(c) + " - " + CellUtils.getErrorMessage(cv.getErrorValue());
                cellError(msg);
                break;
            default:
                msg = "Unexpected cell type detected for cell " + CellUtils.formatAsString(c) + "!";
                cellError(msg);
        }
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
        Iterator<Cell> jt = cells.iterator();
        DataFormatter fmt = new DataFormatter();
        int counter = 0;
        while(it.hasNext()) {
           CellValue cv = it.next();
           Cell cell = jt.next();
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
                                       colValues.add(Workbook.dateTimeFormatter.parse(cv.getStringValue(), dateTimeFormat));
                                    } catch(Exception e) {
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
                           case Numeric:
                               // format according to Excel format
                               colValues.add(fmt.formatCellValue(cell));
                               break;
                           case DateTime:
                               // format according to dateTimeFormatter
                               colValues.add(Workbook.dateTimeFormatter.format(
                                       DateUtil.getJavaDate(cv.getNumberValue()),
                                       dateTimeFormat));
                               break;
                           case String:
                               colValues.add(cv.getStringValue());
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

    public void setDateTimeFormat(String format) {
        this.dateTimeFormat = format;
    }

    public String getDateTimeFormat() {
        return this.dateTimeFormat;
    }
}
