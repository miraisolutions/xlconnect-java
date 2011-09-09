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

package com.miraisolutions.xlconnect;

import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import com.miraisolutions.xlconnect.utils.CellUtils;
import java.io.*;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Class representing a Microsoft Excel Workbook for XLConnect
 * 
 * @author Martin Studer, Mirai Solutions GmbH
 */
public final class Workbook extends Common {
    
    // Logger
    private final static Logger logger = Logger.getLogger("com.miraisolutions.xlconnect");

    // Prefix
    private final static String HEADER = "Header";
    private final static String COLUMN = "Column";
    private final static String SEP = ".";

    // Default style names
    private final static String XLCONNECT_STYLE = "XLCONNECT_STYLE";

    private final static String XLCONNECT_HEADER_STYLE_NAME = "XLConnect.Header";
    private final static String XLCONNECT_GENERAL_STYLE_NAME = "XLConnect.General";
    private final static String XLCONNECT_DATE_STYLE_NAME = "XLConnect.Date";

    // Style types
    private final static String HEADER_STYLE = "Header";
    private final static String NUMERIC_STYLE = "Numeric";
    private final static String STRING_STYLE = "String";
    private final static String BOOLEAN_STYLE = "Boolean";
    private final static String DATETIME_STYLE = "DateTime";

    // Apache POI workbook instance
    private final org.apache.poi.ss.usermodel.Workbook workbook;
    // Underlying file instance
    private File excelFile;
    // Style action
    private StyleAction styleAction = StyleAction.XLCONNECT;
    // Style name prefix
    private String styleNamePrefix = null;
    // Missing value string (null means blank/empty cell)
    private String missingValue = null;
    // Cell style map
    private final Map<String, Map<String, CellStyle>> stylesMap =
            new HashMap<String, Map<String, CellStyle>>(10);
    // Data format map
    private final Map<DataType, String> dataFormatMap = new HashMap<DataType, String>(DataType.values().length);

    // Behavior when detecting an error cell
    // WARN means returning a missing value and registering a warning
    private ErrorBehavior onErrorCell = ErrorBehavior.WARN;


    private Workbook(InputStream in) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(in);
        this.excelFile = null;
        initDefaultDataFormats();
        initDefaultStyles();
    }

    private Workbook(File excelFile) throws FileNotFoundException, IOException, InvalidFormatException {
        this(new FileInputStream(excelFile));
        this.excelFile = excelFile;
    }
    
    private Workbook(File excelFile, SpreadsheetVersion version) {
        switch(version) {
            case EXCEL97:
                this.workbook = new HSSFWorkbook();
                break;
            case EXCEL2007:
                this.workbook = new XSSFWorkbook();
                break;
            default:
                logger.log(Level.SEVERE, "File '" + excelFile.getName() + "': Spreadsheet version not supported!");
                throw new IllegalArgumentException("Spreadsheet version not supported!");
        }

        this.excelFile = excelFile;
        initDefaultDataFormats();
        initDefaultStyles();
    }

    private void initDefaultStyles() {
        Map<String, CellStyle> xlconnectDefaults =
                new HashMap<String, CellStyle>(5);

        // Header style
        CellStyle headerStyle = getCellStyle(XLCONNECT_HEADER_STYLE_NAME);
        if(headerStyle == null) {
            headerStyle = createCellStyle(XLCONNECT_HEADER_STYLE_NAME);
            headerStyle.setDataFormat(dataFormatMap.get(DataType.String));
            headerStyle.setFillPattern(org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setWrapText(true);
        }

        // String / boolean / numeric style
        CellStyle style = getCellStyle(XLCONNECT_GENERAL_STYLE_NAME);
        if(style == null) {
            style = createCellStyle(XLCONNECT_GENERAL_STYLE_NAME);
            style.setDataFormat(dataFormatMap.get(DataType.String));
            style.setWrapText(true);
        }

        // Date style
        CellStyle dateStyle = getCellStyle(XLCONNECT_DATE_STYLE_NAME);
        if(dateStyle == null) {
            dateStyle = createCellStyle(XLCONNECT_DATE_STYLE_NAME);
            dateStyle.setDataFormat(dataFormatMap.get(DataType.DateTime));
            dateStyle.setWrapText(true);
        }

        xlconnectDefaults.put(HEADER_STYLE, headerStyle);
        xlconnectDefaults.put(STRING_STYLE, style);
        xlconnectDefaults.put(NUMERIC_STYLE, style);
        xlconnectDefaults.put(BOOLEAN_STYLE, style);
        xlconnectDefaults.put(DATETIME_STYLE, dateStyle);

        // Add style definitions to style map
        stylesMap.put(XLCONNECT_STYLE, xlconnectDefaults);
    }

    private void initDefaultDataFormats() {
        dataFormatMap.put(DataType.Boolean, "General");
        dataFormatMap.put(DataType.DateTime, "mm/dd/yyyy hh:mm:ss");
        dataFormatMap.put(DataType.Numeric, "General");
        dataFormatMap.put(DataType.String, "General");
    }

    public void setDataFormat(DataType type, String format) {
        logger.log(Level.INFO, "Setting data format '" + format + "' for type " + type.toString());
        dataFormatMap.put(type, format);
    }

    public String getDataFormat(DataType type) {
        return dataFormatMap.get(type);
    }

    public StyleAction getStyleAction() {
        return styleAction;
    }

    public void setStyleAction(StyleAction styleAction) {
        logger.log(Level.INFO, "Setting style action to '" + styleAction.toString() + "'");
        this.styleAction = styleAction;
    }

    public String getStyleNamePrefix() {
        return styleNamePrefix;
    }

    public void setStyleNamePrefix(String styleNamePrefix) {
        logger.log(Level.INFO, "Setting style name prefix to '" + styleNamePrefix + "'");
        this.styleNamePrefix = styleNamePrefix;
    }
    
    public String[] getSheets() {
        int count = workbook.getNumberOfSheets();
        String[] sheetNames = new String[count];

        for(int i = 0; i < count; i++) {
            logger.log(Level.FINE, "Found worksheet '" + workbook.getSheetName(i) + "'");
            sheetNames[i] = workbook.getSheetName(i);
        }

        return sheetNames;
    }

    public int getSheetPos(String sheetName) {
        return workbook.getSheetIndex(sheetName);
    }

    public void setSheetPos(String sheetName, int pos) {
        logger.log(Level.INFO, "Setting sheet '" + sheetName + "' to position " + pos);
        workbook.setSheetOrder(sheetName, pos);
    }

    public String[] getDefinedNames(boolean validOnly) {
        int count = workbook.getNumberOfNames();
        // String[] nameNames = new String[count];
        ArrayList<String> nameNames = new ArrayList<String>();

        for(int i = 0; i < count; i++) {
            Name namedRegion = workbook.getNameAt(i);
            // if valid only, check corresponding reference formula validity
            if(validOnly && !isValidReference(namedRegion.getRefersToFormula())) continue;

            logger.log(Level.FINE, "Found name '" + namedRegion.getNameName() + "'");
            nameNames.add(namedRegion.getNameName());
        }

        return nameNames.toArray(new String[nameNames.size()]);
    }

    public boolean existsSheet(String name) {
        return workbook.getSheet(name) != null;
    }

    public boolean existsName(String name) {
        return workbook.getName(name) != null;
    }

    public void createSheet(String name) {
        if(name.length() > 31) {
            logger.log(Level.SEVERE, "Sheet names are not allowed to contain more than 31 characters!");
            throw new IllegalArgumentException("Sheet names are not allowed to contain more than 31 characters!");
        }

        if(workbook.getSheetIndex(name) < 0) {
            logger.log(Level.INFO, "Creating non-existing sheet '" + name + "'");
            workbook.createSheet(name);
        }
    }

    public void removeSheet(int sheetIndex) {
        if(sheetIndex > -1 && sheetIndex < workbook.getNumberOfSheets()) {
            setAlternativeActiveSheet(sheetIndex);
            logger.log(Level.INFO, "Removing sheet " + sheetIndex);
            workbook.removeSheetAt(sheetIndex);
        }
    }

    public void removeSheet(String name) {
        logger.log(Level.INFO, "Removing sheet '" + name + "'");
        removeSheet(workbook.getSheetIndex(name));
    }

    public void renameSheet(int sheetIndex, String newName) {
        renameSheet(workbook.getSheetName(sheetIndex), newName);
    }

    public void renameSheet(String name, String newName) {
        logger.log(Level.INFO, "Renaming sheet '" + name + "' to '" + newName + "'");
        workbook.setSheetName(workbook.getSheetIndex(name), newName);
    }

    public void cloneSheet(int index, String newName) {
        cloneSheet(workbook.getSheetName(index), newName);
    }

    public void cloneSheet(String name, String newName) {
        logger.log(Level.INFO, "Cloning worksheet '" + name + "' to sheet '" + newName + "'");
        Sheet sheet = workbook.cloneSheet(workbook.getSheetIndex(name));
        workbook.setSheetName(workbook.getSheetIndex(sheet), newName);
    }
    
    public void createName(String name, String formula, boolean overwrite) {

        if(existsName(name)) {
            logger.log(Level.INFO, "Name already exists");
            if(overwrite) {
                // Name already exists but we overwrite --> remove
                logger.log(Level.INFO, "Specified to overwrite name if already existing, " +
                        "therefore remove existing name");
                removeName(name);
            } else {
                // Name already exists but we don't want to overwrite --> error
                logger.log(Level.SEVERE, "Specified name already exists - specified to not overwrite");
                throw new IllegalArgumentException("Specified name '" + name + "' already exists!");
            }
        }

        Name cname = workbook.createName();
        try {
            logger.log(Level.INFO, "Creating name '" + name + "' refering to formula '" + formula + "'");
            cname.setNameName(name);
            cname.setRefersToFormula(formula);
        } catch(Exception e) {
            logger.log(Level.SEVERE, "Failed creating name '" + name + "'. Cleaning up.");
            // --> Clean up (= remove) name
            // Need to set dummy name in order to be able to remove it ...
            String dummyNameName = "XLConnectDummyName";
            cname.setNameName(dummyNameName);
            removeName(dummyNameName);
            throw new IllegalArgumentException(e);
        }
    }

    public void removeName(String name) {
        Name cname = workbook.getName(name);
        if(cname != null) {
            logger.log(Level.INFO, "Removing name '" + name + "'");
            workbook.removeName(name);
        }
    }

    public String getReferenceFormula(String name) {
        return getName(name).getRefersToFormula();
    }

    
    public int[] getReferenceCoordinates(String name) {
        Name cname = getName(name);
        AreaReference aref = new AreaReference(cname.getRefersToFormula());
        // Get upper left corner
	CellReference first = aref.getFirstCell();
        // Get lower right corner
	CellReference last = aref.getLastCell();
	int top = first.getRow();
	int bottom = last.getRow();
	int left = first.getCol();
	int right = last.getCol();
        return new int[]{top,left,bottom,right};
    }

    private void writeData(DataFrame data, Sheet sheet, int startRow, int startCol, boolean header) {
        logger.log(Level.INFO, "Writing data of dimension " + data.rows() + " rows & " + data.columns() + " columns" +
                " to sheet '" + sheet.getSheetName() + "' starting at row " + startRow + " and column " + startCol);

        // Get styles
        Map<String, CellStyle> styles = getStyles(data, sheet, startRow, startCol);

        // Define row & column index variables
        int rowIndex = startRow;
        int colIndex = startCol;

        // In case of column headers ...
        if(header && data.hasColumnHeader()) {
            // For each column write corresponding column name
            for(int i = 0; i < data.columns(); i++) {
                logger.log(Level.FINER, "Writing header '" + data.getColumnName(i) + "'");
                Cell cell = getCell(sheet, rowIndex, colIndex + i);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue(data.getColumnName(i));
                setCellStyle(cell, styles.get(HEADER + i));
            }

            ++rowIndex;
        }

        // For each column of data
        for(int i = 0; i < data.columns(); i++) {
            // Get column style
            CellStyle cs = styles.get(COLUMN + i);
            // Depending on column type ...
            switch(data.getColumnType(i)) {
                case Numeric:
                    logger.log(Level.FINE, "Writing numeric column " + i);
                    Vector<Double> numericValues = data.getColumn(i);
                    for(int j = 0; j < data.rows(); j++) {
                        logger.log(Level.FINER, "Writing row " + j);
                        Cell cell = getCell(sheet, rowIndex + j, colIndex);
                        Double d = numericValues.get(j);
                        if(d == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            setMissing(cell);
                        } else {
                            logger.log(Level.FINEST, "Writing double value '" + d.doubleValue() + "'");
                            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                            cell.setCellValue(d.doubleValue());
                            setCellStyle(cell, cs);
                        }
                    }
                    break;
                case String:
                    logger.log(Level.FINE, "Writing string column " + i);
                    Vector<String> stringValues = data.getColumn(i);
                    for(int j = 0; j < data.rows(); j++) {
                        logger.log(Level.FINER, "Writing row " + j);
                        Cell cell = getCell(sheet, rowIndex + j, colIndex);
                        String s = stringValues.get(j);
                        if(s == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            setMissing(cell);
                        } else {
                            logger.log(Level.FINEST, "Writing string value '" + stringValues.get(j) + "'");
                            cell.setCellType(Cell.CELL_TYPE_STRING);
                            cell.setCellValue(stringValues.get(j));
                            setCellStyle(cell, cs);
                        }
                    }
                    break;
                case Boolean:
                    logger.log(Level.FINE, "Writing boolean column " + i);
                    Vector<Boolean> booleanValues = data.getColumn(i);
                    for(int j = 0; j < data.rows(); j++) {
                        logger.log(Level.FINER, "Writing row " + j);
                        Cell cell = getCell(sheet, rowIndex + j, colIndex);
                        Boolean b = booleanValues.get(j);
                        if(b == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            setMissing(cell);
                        } else {
                            logger.log(Level.FINEST, "Writing boolean value '" + booleanValues.get(j).booleanValue() + "'");
                            cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
                            cell.setCellValue(booleanValues.get(j).booleanValue());
                            setCellStyle(cell, cs);
                        }
                    }
                    break;
                case DateTime:
                    logger.log(Level.FINE, "Writing datetime column " + i);
                    Vector<Date> dateValues = data.getColumn(i);
                    for(int j = 0; j < data.rows(); j++) {
                        logger.log(Level.FINER, "Writing row " + j);
                        Cell cell = getCell(sheet, rowIndex + j, colIndex);
                        Date d = dateValues.get(j);
                        if(d == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            setMissing(cell);
                        } else {
                            logger.log(Level.FINEST, "Writing datetime value '" + dateValues.get(j).toString() + "'");
                            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                            // TODO: date formatting
                            cell.setCellValue(d);
                            setCellStyle(cell, cs);
                        }
                    }
                    break;
                default:
                    logger.log(Level.SEVERE, "Column " + i + ": Unknown column type detected!");
                    throw new IllegalArgumentException("Unknown column type detected!");
            }

            ++colIndex;
        }

        // Force formula recalculation for HSSFSheet
        if(isHSSF()) {
            ((HSSFSheet)sheet).setForceFormulaRecalculation(true);
        }
    }

    private DataFrame readData(Sheet sheet, int startRow, int startCol, int nrows, int ncols, boolean header) {
        logger.log(Level.INFO, "Reading data on sheet '" + sheet.getSheetName() + "', start row = " + startRow +
                ", start column = " + startCol + ", #rows = " + nrows + ", #columns = " + ncols + ", header = " + header);
        
        DataFrame data = new DataFrame();

        // Formula evaluator
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.clearAllCachedResultValues();

        // Loop over columns
        for(int col = 0; col < ncols; col++) {
            logger.log(Level.FINE, "Reading column " + col);

            int colIndex = startCol + col;
            // Determine column header
            String columnHeader = null;
            if(header) {
                Cell cell = getCell(sheet, startRow, colIndex, false);
                // Check if there actually is a cell ...
                if(cell != null) {
                    CellValue cv = evaluator.evaluate(cell);
                    if(cv != null) {
                        columnHeader = cv.getStringValue();
                        logger.log(Level.FINE, "Found column header '" + columnHeader + "'");
                    }
                }
            }
            // If it was specified that there is a header but an empty(/non-existing)
            // cell or cell value is found, then use a default column name
            if(columnHeader == null) {
                columnHeader = "Col" + col;
                logger.log(Level.FINE, "Specified to read column headers but no header found - assuming '" +
                       columnHeader + "'");
            }

            ColumnBuilder cb = new ColumnBuilder(nrows);
            // Loop over rows
            for(int row = header ? 1 : 0; row < nrows; row++) {
                int rowIndex = startRow + row;
                logger.log(Level.FINER, "Reading row index " + rowIndex);

                Cell cell = getCell(sheet, rowIndex, colIndex, false);
                String msg = null;

                // In case the cell does not exist ...
                if(cell == null) {
                    cb.addMissing();
                    continue;
                }

                /*
                 * The following is to handle error cells (before they have been evaluated
                 * to a CellValue) and cells which are formulas but have cached errors.
                 */
                if(
                    cell.getCellType() == Cell.CELL_TYPE_ERROR ||
                    (cell.getCellType() == Cell.CELL_TYPE_FORMULA && cell.getCachedFormulaResultType() == Cell.CELL_TYPE_ERROR)
                ) {
                    msg = "Error detected in cell " + CellUtils.formatAsString(cell) + " - " + CellUtils.getErrorMessage(cell.getErrorCellValue());
                    cb.addError(msg);
                    continue;
                }

                CellValue cv = null;
                // Try to evaluate cell;
                // report an error if this fails
                try {
                    cv = evaluator.evaluate(cell);
                } catch(Exception e) {
                    msg = "Error when trying to evaluate cell " + CellUtils.formatAsString(cell) + " - " + e.getMessage();
                    cb.addError(msg);
                    continue;
                }

                // Not sure if this case should ever happen;
                // let's be sure anyway
                if(cv == null){
                    cb.addMissing();
                    continue;
                }

                // Determine (evaluated) cell data type
                switch(cv.getCellType()) {
                    case Cell.CELL_TYPE_BLANK:
                        logger.log(Level.FINEST, "Blank cell. Cannot determine data type - assuming 'smallest' data type boolean");
                        cb.addMissing();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        logger.log(Level.FINEST, "Found data type boolean");
                        cb.addValue(cv, DataType.Boolean);
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if(DateUtil.isCellDateFormatted(cell)) {
                            logger.log(Level.FINEST, "Found data type datetime");
                            cb.addValue(cv, DataType.DateTime);
                        } else {
                            logger.log(Level.FINEST, "Found data type numeric");
                            cb.addValue(cv, DataType.Numeric);
                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        logger.log(Level.FINEST, "Found data type string");
                        cb.addValue(cv, DataType.String);
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        msg = "Formula detected in already evaluated cell " + CellUtils.formatAsString(cell) + "!";
                        cb.addError(msg);
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        msg = "Error detected in cell " + CellUtils.formatAsString(cell) + " - " + CellUtils.getErrorMessage(cv.getErrorValue());
                        cb.addError(msg);
                        break;
                    default:
                        msg = "Unexpected cell type detected for cell " + CellUtils.formatAsString(cell) + "!";
                        cb.addError(msg);
                }
            }

            // Determine data type for column
            logger.log(Level.FINE, "Determining column type based on row types ...");
            DataType columnType = determineColumnType(cb.detectedTypes);
            switch(columnType) {
                case Boolean:
                {
                    logger.log(Level.FINER, "Determined column " + col + " to be of data type boolean");
                    Vector<Boolean> booleanValues = new Vector(cb.values.size());
                    Iterator<CellValue> it = cb.values.iterator();
                    while(it.hasNext()) {
                        CellValue cv = it.next();
                        if(cv == null)
                        {
                            logger.log(Level.FINEST, "Missing value detected");
                            booleanValues.add(null);
                        } else {
                            logger.log(Level.FINEST, "Reading boolean value '" + cv.getBooleanValue() + "'");
                            booleanValues.add(cv.getBooleanValue());
                        }
                    }
                    data.addColumn(columnHeader, columnType, booleanValues);
                    break;
                }
                case DateTime:
                {
                    logger.log(Level.FINER, "Determined column " + col + " to be of data type datetime");
                    Vector<Date> dateValues = new Vector(cb.values.size());
                    Iterator<CellValue> it = cb.values.iterator();
                    while(it.hasNext()) {
                        CellValue cv = it.next();
                        if(cv == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            dateValues.add(null);
                        } else {
                            logger.log(Level.FINEST, "Reading datetime value '" + DateUtil.getJavaDate(cv.getNumberValue()) + "'");
                            dateValues.add(DateUtil.getJavaDate(cv.getNumberValue()));
                        }
                    }
                    data.addColumn(columnHeader, columnType, dateValues);
                    break;
                }
                case Numeric:
                {
                    logger.log(Level.FINER, "Determined column " + col + " to be of data type numeric");
                    Vector<Double> numericValues = new Vector(cb.values.size());
                    Iterator<CellValue> it = cb.values.iterator();
                    while(it.hasNext()) {
                        CellValue cv = it.next();
                        if(cv == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            numericValues.add(null);
                        } else {
                            Double d = null;
                            switch(cv.getCellType()) {
                                case Cell.CELL_TYPE_BLANK:
                                    break;
                                case Cell.CELL_TYPE_BOOLEAN:
                                    d = cv.getBooleanValue() ? 1.0 : 0.0;
                                    break;
                                default:
                                    d = cv.getNumberValue();
                            }
                            logger.log(Level.FINEST, "Reading numeric value '" + d + "'");
                            numericValues.add(d);
                        }
                    }
                    data.addColumn(columnHeader, columnType, numericValues);
                    break;
                }
                case String:
                {
                    logger.log(Level.FINER, "Determined column " + col + " to be of data type string");
                    Vector<String> stringValues = new Vector(cb.values.size());
                    Iterator<CellValue> it = cb.values.iterator();
                    while(it.hasNext()) {
                        CellValue cv = it.next();
                        if(cv == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            stringValues.add(null);
                        } else {
                            String s = null;
                            switch(cv.getCellType()) {
                                case Cell.CELL_TYPE_BLANK:
                                    break;
                                case Cell.CELL_TYPE_BOOLEAN:
                                    s = String.valueOf(cv.getBooleanValue());
                                    break;
                                case Cell.CELL_TYPE_NUMERIC:
                                    if(cb.isDate.contains(cv))
                                        s = DateUtil.getJavaDate(cv.getNumberValue()).toString();
                                    else
                                        s = String.valueOf(cv.getNumberValue());
                                    break;
                                default:
                                    s = cv.getStringValue();
                            }
                            logger.log(Level.FINEST, "Reading string value '" + s + "'");
                            stringValues.add(s);
                        }
                    }
                    data.addColumn(columnHeader, columnType, stringValues);
                    break;
                }
                default:
                    logger.log(Level.SEVERE, "Could not determine column type for column " + col);
                    throw new IllegalArgumentException("Unknown column type detected!");
            }
        }

        return data;
    }

    public void onErrorCell(ErrorBehavior eb) {
        logger.log(Level.INFO, "Setting error cell behavior to " + eb.toString());
        this.onErrorCell = eb;
    }

    public void writeNamedRegion(DataFrame data, String name, boolean header) {
        logger.log(Level.INFO, "Writing named region '" + name + "' ...");
        Name cname = getName(name);
        checkName(cname);

        // Get sheet where name is defined in
        Sheet sheet = workbook.getSheet(cname.getSheetName());
        logger.log(Level.FINE, "Found named region '" + name + "' on sheet '" + sheet.getSheetName() + "'");
        logger.log(Level.FINE, "Named region refers to formula '" + cname.getRefersToFormula() + "'");

        AreaReference aref = new AreaReference(cname.getRefersToFormula());
        // Get upper left corner
        CellReference topLeft = aref.getFirstCell();

        // Compute bottom right cell coordinates
        int bottomRightRow = topLeft.getRow() + data.rows() - 1;
        if(header) ++bottomRightRow;
        int bottomRightCol = topLeft.getCol() + data.columns() - 1;
        // Create bottom right cell reference
        CellReference bottomRight = new CellReference(sheet.getSheetName(), bottomRightRow,
                bottomRightCol, true, true);

        // Define named range area
        aref = new AreaReference(topLeft, bottomRight);
        // Redefine named range
        cname.setRefersToFormula(aref.formatAsString());

        writeData(data, sheet, topLeft.getRow(), topLeft.getCol(), header);
    }

    public DataFrame readNamedRegion(String name, boolean header) {
        logger.log(Level.INFO, "Reading named region '" + name + "' ... (header = " + header + ")");
        Name cname = getName(name);
        checkName(cname);

        // Get sheet where name is defined in
        Sheet sheet = workbook.getSheet(cname.getSheetName());
        logger.log(Level.FINE, "Found named region '" + name + "' on sheet '" + sheet.getSheetName() + "'");
        logger.log(Level.FINE, "Named region refers to formula '" + cname.getRefersToFormula() + "'");

        AreaReference aref = new AreaReference(cname.getRefersToFormula());
        // Get name corners (top left, bottom right)
        CellReference topLeft = aref.getFirstCell();
        CellReference bottomRight = aref.getLastCell();

        // Determine number of rows and columns
        int nrows = bottomRight.getRow() - topLeft.getRow() + 1;
        int ncols = bottomRight.getCol() - topLeft.getCol() + 1;

        return readData(sheet, topLeft.getRow(), topLeft.getCol(), nrows, ncols, header);
    }

    /**
     * Writes a data frame into the specified worksheet index at the specified location
     *
     * @param data              Data frame to be written to the worksheet
     * @param worksheetIndex    Worksheet index (0-based)
     * @param startRow          Start row (row index of top left cell)
     * @param startCol          Start column (column index of top left cell)
     * @param header            If true, column headers are written, otherwise not
     */
    public void writeWorksheet(DataFrame data, int worksheetIndex, int startRow, int startCol, boolean header) {
        logger.log(Level.INFO, "Writing data to worksheet index " + worksheetIndex + ", start row = " + startRow +
                ", start column = " + startCol);
        Sheet sheet = workbook.getSheetAt(worksheetIndex);
        writeData(data, sheet, startRow, startCol, header);
    }

    public void writeWorksheet(DataFrame data, String worksheetName, int startRow, int startCol, boolean header) {
        logger.log(Level.INFO, "Writing data to worksheet '" + worksheetName + "'");
        writeWorksheet(data, workbook.getSheetIndex(worksheetName), startRow, startCol, header);
    }

    public void writeWorksheet(DataFrame data, int worksheetIndex, boolean header) {
        writeWorksheet(data, worksheetIndex, 0, 0, header);
    }

    public void writeWorksheet(DataFrame data, String worksheetName, boolean header) {
        writeWorksheet(data, worksheetName, 0, 0, header);
    }

    /**
     * Reads data from a worksheet. Data regions can be narrowed down by specifying corresponding row and column ranges.
     * Limits specified as negative integers will be automatically determined. The rules for automatically determining
     * the ranges are as follows:
     *
     * - If start row < 0: get first row on sheet
     * - If end row < 0: get last row on sheet
     * - If start column < 0: get column of first (non-null) cell in start row
     * - If end column < 0: get max column between start row and end row
     *
     * @param worksheetIndex    Worksheet index
     * @param startRow          Start row
     * @param startCol          Start column
     * @param endRow            End row
     * @param endCol            End column
     * @param header            If true, assume header, otherwise not
     * @return                  Data Frame
     */
    public DataFrame readWorksheet(int worksheetIndex, int startRow, int startCol, int endRow, int endCol, boolean header) {
        logger.log(Level.INFO, "Reading worksheet " + worksheetIndex + ", start row = " + startRow + ", start column = " +
                startCol + ", end row = " + endRow + ", end column = " + endCol + ", header = " + header);
        Sheet sheet = workbook.getSheetAt(worksheetIndex);

        if(startRow < 0) startRow = sheet.getFirstRowNum();
        if(startRow < 0) {
            logger.log(Level.SEVERE, "Start row cannot be determined!");
            throw new IllegalArgumentException("Start row cannot be determined!");
        }

        // Check that the start row actually exists
        if(sheet.getRow(startRow) == null) {
            logger.log(Level.SEVERE, "Specified sheet contains no data!");
            throw new IllegalArgumentException("Specified sheet does not contain any data!");
        }

        if(endRow < 0) endRow = sheet.getLastRowNum();

        if(startCol < 0) {
            startCol = sheet.getRow(startRow).getFirstCellNum();
            for(int i = startRow; i <= endRow; i++) {
                Row r = sheet.getRow(i);
                if(r != null && r.getFirstCellNum() < startCol)
                    startCol = r.getFirstCellNum();
            }
        }
        if(startCol < 0) {
            logger.log(Level.SEVERE, "Start column cannot be determined!");
            throw new IllegalArgumentException("Start column cannot be determined!");
        }
        
        if(endCol < 0) {
            endCol = startCol;
            for(int i = startRow; i <= endRow; i++) {
                Row r = sheet.getRow(i);
                // NOTE: getLastCellNum is 1-based!
                if(r != null && (r.getLastCellNum() - 1) > endCol)
                    endCol = r.getLastCellNum() - 1;
            }
        }

        return readData(sheet, startRow, startCol, (endRow - startRow) + 1, (endCol - startCol) + 1, header);
    }

    public DataFrame readWorksheet(int worksheetIndex, boolean header) {
        return readWorksheet(worksheetIndex, -1, -1, -1, -1, header);
    }

    public DataFrame readWorksheet(String worksheetName, int startRow, int startCol, int endRow, int endCol, boolean header) {
        logger.log(Level.INFO, "Reading worksheet '" + worksheetName + "'");
        return readWorksheet(workbook.getSheetIndex(worksheetName), startRow, startCol, endRow, endCol, header);
    }

    public DataFrame readWorksheet(String worksheetName, boolean header) {
        return readWorksheet(worksheetName, -1, -1, -1, -1, header);
    }

    public void addImage(File imageFile, String name, boolean originalSize) throws FileNotFoundException, IOException {
        logger.log(Level.INFO, "Adding image '" + imageFile.getName() + "', original size = " + originalSize);
        Name cname = getName(name);

        // Get sheet where name is defined in
        Sheet sheet = workbook.getSheet(cname.getSheetName());
        logger.log(Level.FINE, "Found named region '" + name + "' on sheet '" + sheet.getSheetName() + "'");
        logger.log(Level.FINE, "Named region refers to formula '" + cname.getRefersToFormula() + "'");
        
        AreaReference aref = new AreaReference(cname.getRefersToFormula());
        // Get name corners (top left, bottom right)
        CellReference topLeft = aref.getFirstCell();
        CellReference bottomRight = aref.getLastCell();

        // Determine image type
        int imageType;
        String filename = imageFile.getName().toLowerCase();
        if(filename.endsWith("jpg") || filename.endsWith("jpeg")) {
            imageType = org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_JPEG;
        } else if(filename.endsWith("png")) {
            imageType = org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_PNG;
        } else if(filename.endsWith("wmf")) {
            imageType = org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_WMF;
        } else if(filename.endsWith("emf")) {
            imageType = org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_EMF;
        } else if(filename.endsWith("bmp") || filename.endsWith("dib")) {
            imageType = org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_DIB;
        } else if(filename.endsWith("pict") || filename.endsWith("pct") || filename.endsWith("pic")) {
            imageType = org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_PICT;
        } else {
            logger.log(Level.SEVERE, "Image type not supported!");
            throw new IllegalArgumentException("Image type not supported!");
        }

        InputStream is = new FileInputStream(imageFile);
        byte[] bytes = IOUtils.toByteArray(is);
        int imageIndex = workbook.addPicture(bytes, imageType);
        is.close();

        Drawing drawing = null;
        if(isHSSF()) {
            drawing = ((HSSFSheet)sheet).getDrawingPatriarch();
            if(drawing == null) {
                drawing = sheet.createDrawingPatriarch();
            }
        } else if(isXSSF()) {
            drawing = ((XSSFSheet)sheet).createDrawingPatriarch();
        } else {
            drawing = sheet.createDrawingPatriarch();
        }

        CreationHelper helper = workbook.getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setRow1(topLeft.getRow());
        anchor.setCol1(topLeft.getCol());
        // +1 since we want to include the
        anchor.setRow2(bottomRight.getRow() + 1);
        anchor.setCol2(bottomRight.getCol() + 1);
        anchor.setAnchorType(ClientAnchor.DONT_MOVE_AND_RESIZE);

        Picture picture = drawing.createPicture(anchor, imageIndex);
        if(originalSize) picture.resize();
    }

    public void addImage(String filename, String name, boolean originalSize) throws FileNotFoundException, IOException {
        addImage(new File(filename), name, originalSize);
    }

    public CellStyle createCellStyle(String name) {
        if(getCellStyle(name) == null) {
            if(isHSSF()) {
                return HCellStyle.create((HSSFWorkbook) workbook, name);
            } else if(isXSSF()) {
                return XCellStyle.create((XSSFWorkbook) workbook, name);
            }
            return null;
        } else {
            logger.log(Level.SEVERE, "Cell style with name '" + name + "' already exists!");
            throw new IllegalArgumentException("Cell style with name '" + name + "' already exists!");
        }
    }

    public CellStyle createCellStyle() {
        return createCellStyle(null);
    }

    public int getActiveSheetIndex() {
        if(workbook.getNumberOfSheets() < 1)
            return -1;
        else
            return workbook.getActiveSheetIndex();
    }

    public String getActiveSheetName() {
        if(workbook.getNumberOfSheets() < 1)
            return null;
        else
            return workbook.getSheetName(workbook.getActiveSheetIndex());
    }

    public void setActiveSheet(int sheetIndex) {
        logger.log(Level.INFO, "Setting active sheet index: " + sheetIndex);
        workbook.setActiveSheet(sheetIndex);
    }

    public void setActiveSheet(String sheetName) {
        logger.log(Level.INFO, "Setting active sheet: " + sheetName);
        int sheetIndex  = workbook.getSheetIndex(sheetName);
        setActiveSheet(sheetIndex);
    }

    public void hideSheet(int sheetIndex, boolean veryHidden) {
        setAlternativeActiveSheet(sheetIndex);
        logger.log(Level.INFO, (veryHidden ? "Very hiding" : "Hiding") + " sheet with index " + sheetIndex);
        workbook.setSheetHidden(sheetIndex, veryHidden ? 
            org.apache.poi.ss.usermodel.Workbook.SHEET_STATE_VERY_HIDDEN :
            org.apache.poi.ss.usermodel.Workbook.SHEET_STATE_HIDDEN);
    }

    public void hideSheet(String sheetName, boolean veryHidden) {
        logger.log(Level.INFO, (veryHidden ? "Very hiding" : "Hiding") + " sheet '" + sheetName + "'");
        hideSheet(workbook.getSheetIndex(sheetName), veryHidden);
    }

    public void unhideSheet(int sheetIndex) {
        logger.log(Level.INFO, "Unhiding sheet " + sheetIndex);
        workbook.setSheetHidden(sheetIndex, org.apache.poi.ss.usermodel.Workbook.SHEET_STATE_VISIBLE);
    }

    public void unhideSheet(String sheetName) {
        logger.log(Level.INFO, "Unhiding sheet '" + sheetName + "'");
        unhideSheet(workbook.getSheetIndex(sheetName));
    }

    public boolean isSheetHidden(int sheetIndex) {
        return workbook.isSheetHidden(sheetIndex);
    }

    public boolean isSheetHidden(String sheetName) {
        return isSheetHidden(workbook.getSheetIndex(sheetName));
    }

    public boolean isSheetVeryHidden(int sheetIndex) {
        return workbook.isSheetVeryHidden(sheetIndex);
    }

    public boolean isSheetVeryHidden(String sheetName) {
        return isSheetVeryHidden(workbook.getSheetIndex(sheetName));
    }
    
    public void setColumnWidth(int sheetIndex, int columnIndex, int width) {
        logger.log(Level.INFO, "Setting width of column " + columnIndex + " on sheet " +
                sheetIndex + " to " + width + " (in units of 1/256th of a character width)");
        getSheet(sheetIndex).setColumnWidth(columnIndex, width);
    }

    public void setColumnWidth(String sheetName, int columnIndex, int width) {
        setColumnWidth(workbook.getSheetIndex(sheetName), columnIndex, width);
    }

    public void setRowHeight(int sheetIndex, int rowIndex, float height) {
        Row r = getSheet(sheetIndex).getRow(rowIndex);
        if(r == null) {
            logger.log(Level.INFO, "Row does not exist - creating it.");
            r = getSheet(sheetIndex).createRow(rowIndex);
        }

        logger.log(Level.INFO, "Setting row " + rowIndex + " of sheet " + sheetIndex +
                " to height " + height + " (in points)");
        r.setHeightInPoints(height);
    }

    public void setRowHeight(String sheetName, int rowIndex, float height) {
        setRowHeight(workbook.getSheetIndex(sheetName), rowIndex, height);
    }

    public void save(File f) throws FileNotFoundException, IOException {
        logger.log(Level.INFO, "Saving workbook to '" + f.getCanonicalPath() + "'");
        FileOutputStream fos = new FileOutputStream(f, false);
        workbook.write(fos);
        fos.close();
    }

    public void save(String file) throws FileNotFoundException, IOException {
        save(new File(file));
    }

    public void save() throws FileNotFoundException, IOException {
        save(excelFile);
    }

    private Name getName(String name) {
        Name cname = workbook.getName(name);
        if(cname != null)
            return cname;
        else
            logger.log(Level.SEVERE, "Name '" + name + "' does not exist!");
            throw new IllegalArgumentException("Name '" + name + "' does not exist!");
    }

    // Checks only if the reference as such is valid
    private boolean isValidReference(String reference) {
        return reference != null && !reference.startsWith("#REF!") && !reference.startsWith("#NULL!");
    }

    private void checkName(Name name) {
        if(!isValidReference(name.getRefersToFormula())) {
            logger.log(Level.SEVERE, "Name '" + name.getNameName() + "' has invalid reference!");
            throw new IllegalArgumentException("Name '" + name.getNameName() + "' has invalid reference!");
        }
        else if(!existsSheet(name.getSheetName())) {
            // The reference as such is valid but it doesn't point to a (existing) sheet ...
            logger.log(Level.SEVERE, "Name '" + name.getNameName() + "' does not refer to a valid sheet!");
            throw new IllegalArgumentException("Name '" + name.getNameName() + "' does not refer to a valid sheet!");
        }
    }

    private boolean isXSSF() {
        return workbook instanceof XSSFWorkbook;
    }

    private boolean isHSSF() {
        return workbook instanceof HSSFWorkbook;
    }

    private Cell getCell(Sheet sheet, int rowIndex, int colIndex, boolean create) {
        // Get or create row
        Row row = sheet.getRow(rowIndex);
        if(row == null) {
            if(create) {
                row = sheet.createRow(rowIndex);
            }
            else return null;
        }
        // Get or create cell
        Cell cell = row.getCell(colIndex);
        if(cell == null) {
            if(create) {
                cell = row.createCell(colIndex);
            }
            else return null;
        }

        return cell;
    }

    private Cell getCell(Sheet sheet, int rowIndex, int colIndex) {
        return getCell(sheet, rowIndex, colIndex, true);
    }

    private Sheet getSheet(int sheetIndex) {
        if(sheetIndex < 0 || sheetIndex >= workbook.getNumberOfSheets()) {
            String msg = "Sheet with index " + sheetIndex + " does not exist!";
            logger.log(Level.SEVERE, msg);
            throw new IllegalArgumentException(msg);
        }
        return workbook.getSheetAt(sheetIndex);
    }

    private Sheet getSheet(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if(sheet == null) {
            String msg = "Sheet with name '" + sheetName + "' does not exist!";
            logger.log(Level.SEVERE, msg);
            throw new IllegalArgumentException(msg);
        }
        return sheet;
    }

    public void setMissingValue(String value) {
        missingValue = value;
    }

    private void setMissing(Cell cell) {
        if(missingValue == null)
            cell.setCellType(Cell.CELL_TYPE_BLANK);
        else {
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(missingValue);
            setCellStyle(cell, DataFormatOnlyCellStyle.get());
        }
    }

    private DataType determineColumnType(Vector<DataType> cellTypes) {
        DataType columnType = DataType.Boolean;

        // Iterate over cell types; as soon as String is detecte we can stop
        Iterator<DataType> it = cellTypes.iterator();
        while(it.hasNext() && !columnType.equals(DataType.String)) {
            DataType dt = it.next();
            // In case current data type ordinal is bigger than column data type ordinal
            // then adapt column data type to be current data type;
            // this assumes DataType enum to in order from "smallest" to "biggest" data type
            if(dt.ordinal() > columnType.ordinal()) columnType = dt;
        }

        return columnType;
    }

    /**
     * Function to set an alternative active sheet in the case
     * the sheet to hide or remove is the currently active sheet
     * in the workbook.
     * If this would not be done, strange behaviour could result
     * when opening an Excel file.
     *
     * @param sheetIndex Sheet to hide or remove
     * @throws IllegalArgumentException In case no alternative active sheet can be found
     */
    private void setAlternativeActiveSheet(int sheetIndex) {
        if(sheetIndex == getActiveSheetIndex()) {
            logger.log(Level.INFO, "Sheet to hide or remove is the currently active sheet in the workbook. " +
                    "Relocating active sheet.");
            // Set active sheet to be first non-hidden/non-very-hidden sheet
            // in the workbook; if there are no such sheets left,
            // then throw an exception
            boolean ok = false;
            for(int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if(i != sheetIndex && !workbook.isSheetHidden(i) && !workbook.isSheetVeryHidden(i)) {
                    setActiveSheet(i);
                    ok = true;
                    break;
                }
            }

            if(!ok) throw new IllegalArgumentException("Cannot hide or remove sheet as there would be no " +
                    "alternative active sheet left!");
        }
    }

    /**
     * Gets a cell style by name.
     *
     * @param name  Cell style name
     * @return      The corresponding cell style if there exists one with the specified name;
     *              null otherwise
     */
    public CellStyle getCellStyle(String name) {
        if(isHSSF()) {
            return HCellStyle.get((HSSFWorkbook) workbook, name);
        } else if(isXSSF()) {
            return XCellStyle.get((XSSFWorkbook) workbook, name);
        }      
        return null;
    }

    private CellStyle getCellStyle(Cell cell) {
        return new SSCellStyle(workbook, cell.getCellStyle());
    }

    public void setCellStyle(Cell c, CellStyle cs) {
        if(cs != null) {
            if(cs instanceof HCellStyle) {
                HCellStyle.set((HSSFCell) c, (HCellStyle) cs);
            } else if(cs instanceof XCellStyle) {
                XCellStyle.set((XSSFCell) c, (XCellStyle) cs);
            } else if(cs instanceof DataFormatOnlyCellStyle) {
                CellStyle csx = getCellStyle(c);
                switch(c.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        if(DateUtil.isCellDateFormatted(c))
                            csx.setDataFormat(dataFormatMap.get(DataType.DateTime));
                        else
                            csx.setDataFormat(dataFormatMap.get(DataType.Numeric));
                        break;
                    case Cell.CELL_TYPE_STRING:
                        csx.setDataFormat(dataFormatMap.get(DataType.String));
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        csx.setDataFormat(dataFormatMap.get(DataType.Boolean));
                        break;
                    default:
                        logger.log(Level.SEVERE, "Unexpected cell type detected!");
                        throw new IllegalArgumentException("Unexpected cell type detected!");
                }
                SSCellStyle.set(c, (SSCellStyle) csx);
            } else {
                SSCellStyle.set(c, (SSCellStyle) cs);
            }
        }
    }
    
    public void setCellStyle(String formula, CellStyle cs) {
        AreaReference aref = new AreaReference(formula);
        String sheetName = aref.getFirstCell().getSheetName();
        Sheet sheet = getSheet(sheetName);
        
        CellReference[] crefs = aref.getAllReferencedCells();
        for(CellReference cref : crefs) {
            Cell c = getCell(sheet, cref.getRow(), cref.getCol());
            setCellStyle(c, cs);
        }
    }

    public void setCellStyle(int sheetIndex, int row, int col, CellStyle cs) {
        Cell c = getCell(getSheet(sheetIndex), row, col);
        setCellStyle(c, cs);
    }

    public void setCellStyle(String sheetName, int row, int col, CellStyle cs) {
        Cell c = getCell(getSheet(sheetName), row, col);
        setCellStyle(c, cs);
    }

    /**
     * Determines the cell styles for headers and columns by column based on the defined style action.
     *
     * @param data      Data frame to be written
     * @param sheet     Worksheet
     * @param startRow  Start row in specified sheet for beginning to write the specified data frame
     * @param startCol  Start column in specified sheet for beginning to write the specified data frame
     * @return          A mapping of header/column indices to cell styles
     */
    private Map<String, CellStyle> getStyles(DataFrame data, Sheet sheet, int startRow, int startCol) {
        Map<String, CellStyle> cstyles = new HashMap<String, CellStyle>(data.columns());

        switch(styleAction) {
            case XLCONNECT:
                Map<String, CellStyle> xlconnectStyles = stylesMap.get(XLCONNECT_STYLE);
                if(data.hasColumnHeader()) {
                    for(int i = 0; i < data.columns(); i++)
                        cstyles.put(HEADER + i, xlconnectStyles.get(HEADER_STYLE));
                }
                for(int i = 0; i < data.columns(); i++) {
                    switch(data.getColumnType(i)) {
                        case Boolean:
                            cstyles.put(COLUMN + i, xlconnectStyles.get(BOOLEAN_STYLE));
                            break;
                        case DateTime:
                            cstyles.put(COLUMN + i, xlconnectStyles.get(DATETIME_STYLE));
                            break;
                        case Numeric:
                            cstyles.put(COLUMN + i, xlconnectStyles.get(NUMERIC_STYLE));
                            break;
                        case String:
                            cstyles.put(COLUMN + i, xlconnectStyles.get(STRING_STYLE));
                            break;
                        default:
                            logger.log(Level.SEVERE, "Unknown column type detected!");
                            throw new IllegalArgumentException("Unknown column type detected!");
                    }
                }
                break;
            case NONE:
                break;
            case PREDEFINED:
                // In case of a header, determine header styles
                if(data.hasColumnHeader()) {
                    for(int i = 0; i < data.columns(); i++) {
                        cstyles.put(HEADER + i, getCellStyle(getCell(sheet, startRow, startCol + i)));
                    }
                }
                int styleRow = startRow + (data.hasColumnHeader() ? 1 : 0);
                for(int i = 0; i < data.columns(); i++) {
                    Cell cell = getCell(sheet, styleRow, startCol + i);
                    cstyles.put(COLUMN + i, getCellStyle(cell));
                }
                break;
            case STYLE_NAME_PREFIX:
                if(data.hasColumnHeader()) {
                    for(int i = 0; i < data.columns(); i++) {
                        String prefix = styleNamePrefix + SEP + HEADER;
                        // Check for style <STYLE_NAME_PREFIX><SEP><HEADER><SEP><COLUMN_NAME>
                        CellStyle cs = getCellStyle(prefix + SEP + data.getColumnName(i));
                        // Check for style <STYLE_NAME_PREFIX><SEP><HEADER><SEP><COLUMN_INDEX>
                        if(cs == null) {
                            logger.log(Level.INFO, "No header style for column '" + data.getColumnName(i) +
                                    "' (specified by column name) found.");
                            cs = getCellStyle(prefix + SEP + (i + 1));
                        }
                        // Check for style <STYLE_NAME_PREFIX><SEP><HEADER>
                        if(cs == null) {
                            logger.log(Level.INFO, "No header style for column '" + data.getColumnName(i) +
                                    "' (specified by index) found.");
                            cs = getCellStyle(prefix);
                        }
                        if(cs == null) {
                            logger.log(Level.WARNING, "No header style found for header '" +
                                    data.getColumnName(i) + "' - taking default");
                            cs = new SSCellStyle(workbook, workbook.getCellStyleAt((short)0));
                        }
                        
                        cstyles.put(HEADER + i, cs);
                    }
                }
                for(int i = 0; i < data.columns(); i++) {
                    String prefix = styleNamePrefix + SEP + COLUMN;
                    // Check for style <STYLE_NAME_PREFIX><SEP><COLUMN><SEP><COLUMN_NAME>
                    CellStyle cs = getCellStyle(prefix + SEP + data.getColumnName(i));
                    // Check for style <STYLE_NAME_PREFIX><SEP><COLUMN><SEP><COLUMN_INDEX>
                    if(cs == null) {
                        logger.log(Level.INFO, "No column style for column '" + data.getColumnName(i) +
                                "' (specified by column name) found.");
                        cs = getCellStyle(prefix + SEP + (i + 1));
                    }
                    // Check for style <STYLE_NAME_PREFIX><SEP><COLUMN><SEP><DATA_TYPE>
                    if(cs == null) {
                        logger.log(Level.INFO, "No column style for column '" + data.getColumnName(i) +
                                "' (specified by index) found.");
                        cs = getCellStyle(prefix + SEP + data.getColumnType(i).toString());
                    }
                    if(cs == null) {
                        logger.log(Level.WARNING, "No column style found for column '" +
                                data.getColumnName(i) + "' - taking default");
                        cs =  new SSCellStyle(workbook, workbook.getCellStyleAt((short)0));
                    }

                    cstyles.put(COLUMN + i, cs);
                }
                break;
            case DATA_FORMAT_ONLY:
                CellStyle cs = DataFormatOnlyCellStyle.get();
                if(data.hasColumnHeader()) {
                    for(int i = 0; i < data.columns(); i++) {
                        cstyles.put(HEADER + i, cs);
                    }
                }
                for(int i = 0; i < data.columns(); i++) {
                    cstyles.put(COLUMN + i, cs);
                }
                break;
            default:
                logger.log(Level.SEVERE, "Style action not supported!");
                throw new IllegalArgumentException("Style action not supported!");
        }

        return cstyles;
    }

    public void mergeCells(int sheetIndex, String reference) {
        getSheet(sheetIndex).addMergedRegion(CellRangeAddress.valueOf(reference));
    }

    public void mergeCells(String sheetName, String reference) {
        getSheet(sheetName).addMergedRegion(CellRangeAddress.valueOf(reference));
    }

    public void unmergeCells(int sheetIndex, String reference) {
        Sheet sheet = getSheet(sheetIndex);
        for(int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress cra = sheet.getMergedRegion(i);
            if(cra.formatAsString().equals(reference)) {
                sheet.removeMergedRegion(i);
                break;
            }
        }
    }

    public void unmergeCells(String sheetName, String reference) {
        unmergeCells(workbook.getSheetIndex(sheetName), reference);
    }
    

    /**
     * Get the workbook from a Microsoft Excel file.
     *
     * Reads the workbook if the file exists, otherwise creates a new workbook of the corresponding format.
     *
     * @param excelfile Microsoft Excel file to read or create if not existing
     * @return Instance of the workbook
     * @throws FileNotFoundException
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static Workbook getWorkbook(File excelFile, boolean create) throws FileNotFoundException, IOException, InvalidFormatException {
        Workbook wb;

        if(excelFile.exists()) {
            logger.log(Level.INFO, "Creating XLConnect workbook instance for existing file '" + excelFile.getCanonicalPath() + "'");
            wb = new Workbook(excelFile);
        } else {
            if(create) {
                logger.log(Level.INFO, "Creating XLConnect workbook instance for new file '" + excelFile.getCanonicalPath() + "'");
                String filename = excelFile.getName().toLowerCase();
                if(filename.endsWith(".xls")) {
                    wb = new Workbook(excelFile, SpreadsheetVersion.EXCEL97);
                } else if(filename.endsWith(".xlsx")) {
                    wb = new Workbook(excelFile, SpreadsheetVersion.EXCEL2007);
                } else {
                    logger.log(Level.SEVERE, "File extension not supported! Only *.xls and *.xlsx are allowed!");
                    throw new IllegalArgumentException("File extension not supported! Only *.xls and *.xlsx are allowed!");
                }
            } else {
                logger.log(Level.SEVERE, "File '" + excelFile.getName() + "' could not be found - " +
                        "you may specify to automatically create the file if not existing.");
                throw new FileNotFoundException("File '" + excelFile.getName() + "' could not be found - " +
                        "you may specify to automatically create the file if not existing.");
            }
        }

        logger.log(Level.INFO, "Excel version: " + (wb.isHSSF() ? SpreadsheetVersion.EXCEL97.toString() : SpreadsheetVersion.EXCEL2007.toString()));
        return wb;
    }

    public static Workbook getWorkbook(String filename, boolean create) throws FileNotFoundException, IOException, InvalidFormatException {
        return Workbook.getWorkbook(new File(filename), create);
    }

    

    /**
     * Column Builder instance used by readData
     */
    public class ColumnBuilder {

        int nrows;

        // Collection to hold detected data types for each value in a column
        // --> will be used to determine actual final data type for column
        Vector<DataType> detectedTypes;
        // Collection to hold actual values
        Vector<CellValue> values;

        // Helper collection to store CellValue's that are dates
        // This is needed as a CellValue doesn't store the information whether it is
        // a date or not - dates are just numerics
        Vector<CellValue> isDate = new Vector<CellValue>();


        public ColumnBuilder(int nrows) {
            this.nrows = nrows;
            this.detectedTypes = new Vector<DataType>(nrows);
            this.values = new Vector<CellValue>(nrows);
        }

        public void addMissing() {
            logger.log(Level.FINEST, "Cannot determine data type - assuming 'smallest' data type boolean");
            // Add "missing" to collection
            values.add(null);
            // assume "smallest" data type
            detectedTypes.add(DataType.Boolean);
        }

        private void addValue(CellValue cv, DataType dt) {
            if(DataType.DateTime.equals(dt)) isDate.add(cv);
            values.add(cv);
            detectedTypes.add(dt);
        }

        private void addError(String msg) {
            if(onErrorCell.equals(ErrorBehavior.WARN)) {
                logger.log(Level.WARNING, msg);
                values.add(null);
                detectedTypes.add(DataType.Boolean);
                addWarning(msg);
            } else {
                logger.log(Level.SEVERE, msg);
                throw new IllegalArgumentException(msg);
            }
        }
    }

    public void setCellFormula(Cell c, String formula) {
        c.setCellFormula(formula);
    }
    
    public void setCellFormula(String formulaDest, String formulaString) {
        AreaReference aref = new AreaReference(formulaDest);
        String sheetName = aref.getFirstCell().getSheetName();
        Sheet sheet = getSheet(sheetName);
        
        CellReference[] crefs = aref.getAllReferencedCells();
        for(CellReference cref : crefs) {
            Cell c = getCell(sheet, cref.getRow(), cref.getCol());
            setCellFormula(c, formulaString);
        }
    }

    public void setCellFormula(int sheetIndex, int row, int col, String formula) {
        Cell c = getCell(getSheet(sheetIndex), row, col);
        setCellFormula(c, formula);
    }

    public void setCellFormula(String sheetName, int row, int col, String formula) {
        Cell c = getCell(getSheet(sheetName), row, col);
        setCellFormula(c, formula);
    }

    public String getCellFormula(Cell c) {
        return c.getCellFormula();
    }
    
    public String getCellFormula(int sheetIndex, int row, int col) {
        Cell c = getCell(getSheet(sheetIndex), row, col);
        return getCellFormula(c);
    }

    public String getCellFormula(String sheetName, int row, int col) {
        Cell c = getCell(getSheet(sheetName), row, col);
        return getCellFormula(c);
    }

    public boolean getForceFormulaRecalculation(int sheetIndex) { 
        return getSheet(sheetIndex).getForceFormulaRecalculation();
    }

    public boolean getForceFormulaRecalculation(String sheetName) { 
        return getSheet(sheetName).getForceFormulaRecalculation();
    }

    public void setForceFormulaRecalculation(int sheetIndex, boolean value) { 
        getSheet(sheetIndex).setForceFormulaRecalculation(value);
    }

    public void setForceFormulaRecalculation(String sheetName, boolean value) { 
        getSheet(sheetName).setForceFormulaRecalculation(value);
    }

    public void setAutoFilter(int sheetIndex, String reference) {
        getSheet(sheetIndex).setAutoFilter(CellRangeAddress.valueOf(reference));
    }

    public void setAutoFilter(String sheetName, String reference) {
        getSheet(sheetName).setAutoFilter(CellRangeAddress.valueOf(reference));
    }
}
