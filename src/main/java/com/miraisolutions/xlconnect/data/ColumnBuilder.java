/*
 *
    XLConnect
    Copyright (C) 2010-2025 Mirai Solutions GmbH

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

import com.miraisolutions.xlconnect.ErrorBehavior;
import com.miraisolutions.xlconnect.Workbook;
import com.miraisolutions.xlconnect.utils.CellUtils;
import com.zaxxer.sparsebits.SparseBitSet;
import org.apache.poi.ss.usermodel.*;

import java.util.*;

public abstract class ColumnBuilder {

    // Collection to hold detected data types for each value in a column
    // --> will be used to determine actual final data type for column
    protected final ArrayList<DataType> detectedTypes;
    // Collection to hold cell references
    protected final ArrayList<Cell> cells;
    // Collection to hold actual values
    protected final ArrayList<CellValue> values;
    // Date/time format used for conversion to and from strings
    protected final String dateTimeFormat;

    // Should conversion to a less generic data type be forced?
    protected final boolean forceConversion;

    protected final boolean takeCached;
    protected final FormulaEvaluator evaluator;
    protected final ErrorBehavior onErrorCell;

    // This is used to support the warnings mechanism on the R side
    protected final ArrayList<String> warnings = new ArrayList<>();

    public ColumnBuilder(int nrows, boolean forceConversion,
                         boolean takeCached, FormulaEvaluator evaluator, ErrorBehavior onErrorCell,
                         String dateTimeFormat) {

        this.detectedTypes = new ArrayList<>(nrows);
        this.cells = new ArrayList<>(nrows);
        this.values = new ArrayList<>(nrows);
        this.forceConversion = forceConversion;
        this.evaluator = evaluator;
        this.takeCached = takeCached;
        this.onErrorCell = onErrorCell;
        this.dateTimeFormat = dateTimeFormat;
    }

    public void clear() {
        detectedTypes.clear();
        cells.clear();
        values.clear();
        warnings.clear();
    }

    public void addCell(Cell c) {
        try {
            Optional<CellValue> cellValue = Optional.ofNullable(c).map(this::getCellValue);
            if (cellValue.isPresent()) {
                CellValue cv = cellValue.get();
                if (cv.getCellType() == CellType.ERROR) {
                    cellError("Error detected in cell " + CellUtils.formatAsString(c) + " - " +
                            CellUtils.getErrorMessage(c.getErrorCellValue()));
                } else {
                    handleCell(c, cv);
                }
            } else {
                addMissing();
            }
        } catch (Exception e) {
            cellError("Error when trying to evaluate cell " + CellUtils.formatAsString(c) + " - " + e.getMessage());
        }
    }

    protected void addMissing() {
        // Add "missing" to collection
        values.add(null);
        cells.add(null);
        // assume "smallest" data type
        detectedTypes.add(DataType.Boolean);
    }

    protected void addValue(Cell c, CellValue cv, DataType dt) {
        cells.add(c);
        values.add(cv);
        detectedTypes.add(dt);
    }

    public Column buildBooleanColumn() {
        int size = values.size();
        boolean[] colValues = new boolean[size];
        SparseBitSet missing = new SparseBitSet(size);

        int counter = 0;
        for (CellValue cv : values) {
            if (cv == null) {
                missing.set(counter);
            } else {
                switch (detectedTypes.get(counter)) {
                    case Boolean:
                        colValues[counter] = cv.getBooleanValue();
                        break;
                    case Numeric:
                        colValues[counter] = forceConversion && cv.getNumberValue() > 0;
                        missing.set(counter, !forceConversion);
                        break;
                    case String:
                        colValues[counter] = forceConversion && Boolean.parseBoolean(cv.getStringValue().toLowerCase());
                        missing.set(counter, !forceConversion);
                        break;
                    case DateTime:
                        missing.set(counter);
                        this.warnings.add("Cell " + CellUtils.formatAsString(cells.get(counter)) + " cannot be converted from DateTime to Boolean - returning NA");
                        break;
                    default:
                        throw new IllegalArgumentException("Unknown data type detected!");
                }
            }
            ++counter;
        }
        return new Column(colValues, size, missing, DataType.Boolean);
    }

    public Column buildDateTimeColumn() {
        int size = values.size();
        Date[] colValues = new Date[size];
        SparseBitSet missing = new SparseBitSet(size);

        Iterator<CellValue> it = values.iterator();
        Iterator<Cell> jt = cells.iterator();
        int counter = 0;
        while (it.hasNext()) {
            CellValue cv = it.next();
            Cell cell = jt.next();
            if (cv == null) {
                missing.set(counter);
            } else {
                switch (detectedTypes.get(counter)) {
                    case Boolean:
                        missing.set(counter);
                        this.warnings.add("Cell " + CellUtils.formatAsString(cells.get(counter)) + " cannot be converted from Boolean to DateTime - returning NA");
                        break;
                    case Numeric:
                        if (forceConversion) {
                            if (DateUtil.isValidExcelDate(cv.getNumberValue())) {
                                colValues[counter] = cell.getDateCellValue();
                            } else {
                                missing.set(counter);
                                this.warnings.add("Cell " + CellUtils.formatAsString(cells.get(counter)) + " cannot be converted from Numeric to DateTime - returning NA");
                            }
                        } else {
                            missing.set(counter);
                        }
                        break;
                    case String:
                        if (forceConversion) {
                            try {
                                colValues[counter] = Workbook.dateTimeFormatter.parse(cv.getStringValue(), dateTimeFormat);
                            } catch (Exception e) {
                                missing.set(counter);
                                this.warnings.add("Cell " + CellUtils.formatAsString(cells.get(counter)) + " cannot be converted from " +
                                        "String to DateTime - returning NA - cause: " + e.getClass() + ":" + e.getMessage());
                            }
                        } else {
                            missing.set(counter);
                        }
                        break;
                    case DateTime:
                        colValues[counter] = cell.getDateCellValue();
                        break;
                    default:
                        throw new IllegalArgumentException("Unknown data type detected!");
                }
            }
            ++counter;
        }
        return new Column(colValues, size, missing, DataType.DateTime);
    }

    public Column buildNumericColumn() {
        int size = values.size();
        double[] colValues = new double[size];
        SparseBitSet missing = new SparseBitSet(size);

        int counter = 0;
        for (CellValue cv : values) {
            if (cv == null) {
                missing.set(counter);
            } else {
                switch (detectedTypes.get(counter)) {
                    case Boolean:
                        colValues[counter] = cv.getBooleanValue() ? 1.0 : 0.0;
                        break;
                    case Numeric:
                        colValues[counter] = cv.getNumberValue();
                        break;
                    case String:
                        if (forceConversion) {
                            try {
                                colValues[counter] = Double.parseDouble(cv.getStringValue());
                            } catch (NumberFormatException e) {
                                missing.set(counter);
                                this.warnings.add("Cell " + CellUtils.formatAsString(cells.get(counter)) +
                                        " cannot be converted from String to Numeric - returning NA");
                            }
                        } else {
                            missing.set(counter);
                        }
                        break;
                    case DateTime:
                        if (forceConversion) {
                            colValues[counter] = cv.getNumberValue();
                        } else {
                            missing.set(counter);
                        }
                        break;
                    default:
                        throw new IllegalArgumentException("Unknown data type detected!");
                }
            }
            ++counter;
        }
        return new Column(colValues, size, missing, DataType.Numeric);
    }

    public Column buildStringColumn() {
        int size = values.size();
        String[] colValues = new String[size];
        SparseBitSet missing = new SparseBitSet(size);

        Iterator<CellValue> it = values.iterator();
        Iterator<Cell> jt = cells.iterator();
        DataFormatter fmt = new DataFormatter();
        int counter = 0;
        while (it.hasNext()) {
            CellValue cv = it.next();
            Cell cell = jt.next();
            if (cv == null) {
                missing.set(counter);
            } else {
                switch (detectedTypes.get(counter)) {
                    case Boolean:
                        colValues[counter] = cv.getBooleanValue() ? "true" : "false";
                        break;
                    case Numeric:
                        // format according to Excel format
                        // see also org.apache.poi.ss.usermodel.DataFormatter#formatRawCellContents
                        int formatIndex = cell.getCellStyle().getDataFormat();
                        String formatStr = cell.getCellStyle().getDataFormatString();
                        colValues[counter] = fmt.formatRawCellContents(cv.getNumberValue(), formatIndex, formatStr);
                        break;
                    case DateTime:
                        // format according to dateTimeFormatter
                        colValues[counter] = Workbook.dateTimeFormatter.format(cell.getDateCellValue(), dateTimeFormat);
                        break;
                    case String:
                        colValues[counter] = cv.getStringValue();
                        break;
                    default:
                        throw new IllegalArgumentException("Unknown data type detected!");
                }
            }
            ++counter;
        }
        return new Column(colValues, size, missing, DataType.String);
    }

    protected void cellError(String msg) {
        if (this.onErrorCell.equals(ErrorBehavior.WARN)) {
            this.addMissing();
            this.warnings.add(msg);
        } else {
            throw new IllegalArgumentException(msg);
        }
    }

    public DataType determineColumnType() {
        DataType columnType = DataType.Boolean;
        // Iterate over cell types; as soon as String is detecte we can stop
        Iterator<DataType> it = detectedTypes.iterator();
        while (it.hasNext() && !columnType.equals(DataType.String)) {
            DataType dt = it.next();
            // In case current data type ordinal is bigger than column data type ordinal
            // then adapt column data type to be current data type;
            // this assumes DataType enum to in order from "smallest" to "biggest" data type
            if (dt.ordinal() > columnType.ordinal()) {
                columnType = dt;
            }
        }
        return columnType;
    }

    // extracts the cached value from a cell without re-evaluating
    // the formula. returns null if the cell is blank.
    protected CellValue getCachedCellValue(Cell cell) {
        CellType valueType = cell.getCellType();
        if (valueType == CellType.FORMULA) {
            valueType = cell.getCachedFormulaResultType();
        }
        switch (valueType) {
            case BLANK:
                return null;
            case BOOLEAN:
                if (cell.getBooleanCellValue()) {
                    return CellValue.TRUE;
                } else {
                    return CellValue.FALSE;
                }
            case NUMERIC:
                return new CellValue(cell.getNumericCellValue());
            case STRING:
                return new CellValue(cell.getStringCellValue());
            case ERROR:
                return CellValue.getError(cell.getErrorCellValue());
            default:
                String msg = String.format("Could not extract value from cell with cached value type %s", valueType);
                throw new RuntimeException(msg);
        }
    }

    // extracts the value from a cell by either evaluating it or taking the
    // cached value
    protected CellValue getCellValue(Cell cell) {
        if (this.takeCached) {
            return getCachedCellValue(cell);
        } else {
            return this.evaluator.evaluate(cell);
        }
    }

    protected abstract void handleCell(Cell c, CellValue cv);

    public List<String> getWarnings() {
        return this.warnings;
    }
}