/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.miraisolutions.xlconnect;

import com.miraisolutions.xlconnect.data.DataFrame;
import com.miraisolutions.xlconnect.data.DataType;
import com.miraisolutions.xlconnect.utils.WorkbookUtils;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Microsoft Excel Workbook Entity
 * 
 * @author Martin Studer, Mirai Solutions GmbH
 */
public final class Workbook {
    
    // Logger
    private final static Logger logger = Logger.getLogger("com.miraisolutions.xlconnect");

    // Prefix
    private final static String HEADER = "H";
    private final static String COLUMN = "C";

    // Default style names
    private final static String XLCONNECT_STYLE = "XLCONNECT_STYLE";

    // Style types
    private final static String HEADER_STYLE = "Header";
    private final static String NUMERIC_STYLE = "Numeric";
    private final static String STRING_STYLE = "String";
    private final static String BOOLEAN_STYLE = "Boolean";
    private final static String DATETIME_STYLE = "DateTime";

    // Apache POI workbook instance
    private org.apache.poi.ss.usermodel.Workbook workbook;
    // Underlying file instance
    private File excelFile;
    // Style action
    private StyleAction styleAction = StyleAction.XLCONNECT;
    // Style name
    private String styleName = null;
    // Cell style map
    private Map<String, Map<String, CellStyle>> stylesMap = new HashMap<String, Map<String, CellStyle>>(10);


    /**
     * CONSTRUCTORS
     */

    private Workbook(InputStream in) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(in);
        this.excelFile = null;
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
        initDefaultStyles();
    }

    private void initDefaultStyles() {
        Map<String, CellStyle> xlconnectDefaults = new HashMap<String, CellStyle>(5);
        DataFormat dataFormat = workbook.createDataFormat();

        // Header style
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setDataFormat(dataFormat.getFormat("General"));
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        headerStyle.setWrapText(true);
        // String / boolean / numeric style
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(dataFormat.getFormat("General"));
        style.setWrapText(true);
        // Date style
        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(dataFormat.getFormat("mm/dd/yyyy hh:mm:ss"));
        dateStyle.setWrapText(true);

        xlconnectDefaults.put(HEADER_STYLE, headerStyle);
        xlconnectDefaults.put(STRING_STYLE, style);
        xlconnectDefaults.put(NUMERIC_STYLE, style);
        xlconnectDefaults.put(BOOLEAN_STYLE, style);
        xlconnectDefaults.put(DATETIME_STYLE, dateStyle);

        // Add style definitions to style map
        stylesMap.put(XLCONNECT_STYLE, xlconnectDefaults);
    }

    public StyleAction getStyleAction() {
        return styleAction;
    }

    public void setStyleAction(StyleAction styleAction) {
        this.styleAction = styleAction;
    }

    public String getStyleName() {
        return styleName;
    }

    public void setStyleName(String styleName) {
        this.styleName = styleName;
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

    public String[] getDefinedNames() {
        int count = workbook.getNumberOfNames();
        String[] nameNames = new String[count];

        for(int i = 0; i < count; i++) {
            Name namedRegion = workbook.getNameAt(i);
            logger.log(Level.FINE, "Found name '" + namedRegion.getNameName() + "'");
            nameNames[i] = namedRegion.getNameName();
        }

        return nameNames;
    }

    public boolean isSheetExisting(String name) {
        return workbook.getSheet(name) != null;
    }

    public boolean isNameExisting(String name) {
        return workbook.getName(name) != null;
    }

    public void createSheet(String name) {
        if(workbook.getSheetIndex(name) < 0) {
            logger.log(Level.INFO, "Creating non-existing sheet '" + name + "'");
            workbook.createSheet(name);
        }
    }

    public void removeSheet(String name) {
        Sheet sheet = workbook.getSheet(name);
        if(sheet != null) {
            int index = workbook.getSheetIndex(sheet);
            logger.log(Level.INFO, "Removing sheet '" + name + "'");
            workbook.removeSheetAt(index);
        }
    }
    
    public void createName(String name, String formula) {
        Name cname = workbook.createName();
        logger.log(Level.INFO, "Creating name '" + name + "' refering to formula '" + formula + "'");
        cname.setNameName(name);
        cname.setRefersToFormula(formula);
    }

    public void removeName(String name) {
        Name cname = workbook.getName(name);
        if(cname != null) {
            logger.log(Level.INFO, "Removing name '" + name + "'");
            workbook.removeName(name);
        }
    }

    private void writeData(DataFrame data, Sheet sheet, int startRow, int startCol) {
        logger.log(Level.INFO, "Writing data of dimension " + data.rows() + " rows & " + data.columns() + " columns" +
                " to sheet '" + sheet.getSheetName() + "' starting at row " + startRow + " and column " + startCol);

        // Get styles
        Map<String, CellStyle> styles = getStyles(data, sheet, startRow, startCol);

        // Define row & column index variables
        int rowIndex = startRow;
        int colIndex = startCol;

        // In case of column headers ...
        if(data.hasColumnHeader()) {
            logger.log(Level.FINE, "Column header detected.");
            // For each column write corresponding column name
            for(int i = 0; i < data.columns(); i++) {
                logger.log(Level.FINER, "Writing header '" + data.getColumnName(i) + "'");
                Cell cell = getCell(sheet, rowIndex, colIndex + i);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue(data.getColumnName(i));
                cell.setCellStyle(styles.get(HEADER + i));
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
                            cell.setCellStyle(cs);
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
                            cell.setCellStyle(cs);
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
                            cell.setCellStyle(cs);
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
                            cell.setCellStyle(cs);
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

            // Collection to hold detected data types for each value in a column
            // --> will be used to determine actual final data type for column
            Vector<DataType> detectedTypes = new Vector<DataType>(nrows);
            // Collection to hold actual values
            Vector<CellValue> values = new Vector<CellValue>(nrows);

            // Loop over rows
            for(int row = header ? 1 : 0; row < nrows; row++) {
                int rowIndex = startRow + row;
                logger.log(Level.FINER, "Reading row index " + rowIndex);

                Cell cell = getCell(sheet, rowIndex, colIndex, false);
                if((cell != null) && (cell.getCellType() != Cell.CELL_TYPE_BLANK) &&
                        (cell.getCellType() != Cell.CELL_TYPE_ERROR)) {
                    
                    CellValue cv = evaluator.evaluate(cell);
                    // Add value to collection
                    values.add(cv);
                    // If "empty" value, continue
                    if(cv == null) {
                        logger.log(Level.FINEST, "Cannot determine data type - assuming 'smallest' data type boolean");
                        // We assume Boolean ("smallest" data type)
                        detectedTypes.add(DataType.Boolean);
                        continue;
                    }

                    // Determine cell data type
                    switch(cv.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                        case Cell.CELL_TYPE_BOOLEAN:
                            logger.log(Level.FINEST, "Found data type boolean");
                            detectedTypes.add(DataType.Boolean);
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if(DateUtil.isCellDateFormatted(cell)) {
                                logger.log(Level.FINEST, "Found data type datetime");
                                detectedTypes.add(DataType.DateTime);
                            } else {
                                logger.log(Level.FINEST, "Found data type numeric");
                                detectedTypes.add(DataType.Numeric);
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:
                            logger.log(Level.FINEST, "Found data type string");
                            detectedTypes.add(DataType.String);
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            logger.log(Level.SEVERE, "Formula detected in already evaluated cell!");
                            throw new IllegalArgumentException("Formula detected in already evaluated cell!");
                        case Cell.CELL_TYPE_ERROR:
                            logger.log(Level.SEVERE, "Cell of type ERROR detected! Invalid formula?");
                            throw new IllegalArgumentException("Cell of type ERROR detected! Invalid formula?");
                        default:
                            logger.log(Level.SEVERE, "Unexpected cell type detected!");
                            throw new IllegalArgumentException("Unexpected cell type detected!");
                    }
                } else {
                    logger.log(Level.FINEST, "Cannot determine data type - assuming 'smallest' data type boolean");
                    // Add "missing" to collection
                    values.add(null);
                    // assume "smallest" data type
                    detectedTypes.add(DataType.Boolean);
                }
            }

            // Determine data type for column
            logger.log(Level.FINE, "Determining column type based on row types ...");
            DataType columnType = determineColumnType(detectedTypes);
            switch(columnType) {
                case Boolean:
                {
                    logger.log(Level.FINER, "Determined column " + col + " to be of data type boolean");
                    Vector<Boolean> booleanValues = new Vector(values.size());
                    Iterator<CellValue> it = values.iterator();
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
                    Vector<Date> dateValues = new Vector(values.size());
                    Iterator<CellValue> it = values.iterator();
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
                    Vector<Double> numericValues = new Vector(values.size());
                    Iterator<CellValue> it = values.iterator();
                    while(it.hasNext()) {
                        CellValue cv = it.next();
                        if(cv == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            numericValues.add(null);
                        } else {
                            logger.log(Level.FINEST, "Reading numeric value '" + cv.getNumberValue() + "'");
                            numericValues.add(cv.getNumberValue());
                        }
                    }
                    data.addColumn(columnHeader, columnType, numericValues);
                    break;
                }
                case String:
                {
                    logger.log(Level.FINER, "Determined column " + col + " to be of data type string");
                    Vector<String> stringValues = new Vector(values.size());
                    Iterator<CellValue> it = values.iterator();
                    while(it.hasNext()) {
                        CellValue cv = it.next();
                        if(cv == null) {
                            logger.log(Level.FINEST, "Missing value detected");
                            stringValues.add(null);
                        } else {
                            logger.log(Level.FINEST, "Reading string value '" + cv.getStringValue() + "'");
                            stringValues.add(cv.getStringValue());
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

    public void writeNamedRegion(DataFrame data, String name) {
        logger.log(Level.INFO, "Writing named region '" + name + "' ...");
        Name cname = getName(name);

        // Get sheet where name is defined in
        Sheet sheet = workbook.getSheet(cname.getSheetName());
        logger.log(Level.FINE, "Found named region '" + name + "' on sheet '" + sheet.getSheetName() + "'");
        logger.log(Level.FINE, "Named region refers to formula '" + cname.getRefersToFormula() + "'");

        AreaReference aref = new AreaReference(cname.getRefersToFormula());
        // Get upper left corner
        CellReference topLeft = aref.getFirstCell();

        // Compute bottom right cell coordinates
        int bottomRightRow = topLeft.getRow() + data.rows() - 1;
        if(data.hasColumnHeader()) ++bottomRightRow;
        int bottomRightCol = topLeft.getCol() + data.columns() - 1;
        // Create bottom right cell reference
        CellReference bottomRight = new CellReference(sheet.getSheetName(), bottomRightRow,
                bottomRightCol, true, true);

        // Define named range area
        aref = new AreaReference(topLeft, bottomRight);
        // Redefine named range
        cname.setRefersToFormula(aref.formatAsString());

        writeData(data, sheet, topLeft.getRow(), topLeft.getCol());
    }

    public void writeNamedRegion(DataFrame data, String name, String location, boolean overwrite) {
        logger.log(Level.INFO, "Writing named region '" + name + "' to location '" + location + "'; overwrite = " +
                overwrite + " ...");
        defineOrOverwriteName(name, location, overwrite);
        writeNamedRegion(data, name);
    }

    public DataFrame readNamedRegion(String name, boolean header) {
        logger.log(Level.INFO, "Reading named region '" + name + "' ... (header = " + header + ")");
        Name cname = getName(name);

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
     */
    public void writeWorksheet(DataFrame data, int worksheetIndex, int startRow, int startCol) {
        logger.log(Level.INFO, "Writing data to worksheet index " + worksheetIndex + ", start row = " + startRow +
                ", start column = " + startCol);
        Sheet sheet = workbook.getSheetAt(worksheetIndex);
        writeData(data, sheet, startRow, startCol);
    }

    public void writeWorksheet(DataFrame data, String worksheetName, int startRow, int startCol, boolean create) {
        logger.log(Level.INFO, "Writing data to worksheet '" + worksheetName + "'");
        if(create) createSheet(worksheetName);
        writeWorksheet(data, workbook.getSheetIndex(worksheetName), startRow, startCol);
    }

    public void writeWorksheet(DataFrame data, int worksheetIndex) {
        writeWorksheet(data, worksheetIndex, 0, 0);
    }

    public void writeWorksheet(DataFrame data, String worksheetName) {
        writeWorksheet(data, worksheetName, 0, 0, false);
    }

    /**
     * Reads data from a worksheet. Data regions can be narrowed down by specifying corresponding row and column ranges.
     * Limits specified as negative integers will be automatically determined. The rules for automatically determining
     * the ranges are as follows:
     *
     * - If start row < 0: get first row on sheet
     * - If end row < 0: get last row on sheet
     * - If start column < 0: get column of first (non-null) cell in start row
     * - If end column < 0: get column of last (non-null) cell in end row
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

        if(endRow < 0) endRow = sheet.getLastRowNum();
        if(endRow < 0) {
            logger.log(Level.SEVERE, "End row cannot be determined!");
            throw new IllegalArgumentException("End row cannot be determined!");
        }

        if(startCol < 0) startCol = sheet.getRow(startRow).getFirstCellNum();
        if(startCol < 0) {
            logger.log(Level.SEVERE, "Start column cannot be determined!");
            throw new IllegalArgumentException("Start column cannot be determined!");
        }
        // NOTE: getLastCellNum is 1-based!
        if(endCol < 0) endCol = sheet.getRow(endRow).getLastCellNum() - 1;
        if(endCol < 0) {
            logger.log(Level.SEVERE, "End column cannot be determined!");
            throw new IllegalArgumentException("End column cannot be determined!");
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

    public void addImage(File imageFile, boolean originalSize, String name) throws FileNotFoundException, IOException {
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

    public void addImage(File imageFile, boolean originalSize, String name, String location,
            boolean overwrite) throws FileNotFoundException, IOException {

        defineOrOverwriteName(name, location, overwrite);
        addImage(imageFile, originalSize, name);
    }

    public void addImage(String filename, boolean originalSize, String name) throws FileNotFoundException, IOException {
        addImage(new File(filename), originalSize, name);
    }

    public void addImage(String filename, boolean originalSize, String name, String location,
            boolean overwrite) throws FileNotFoundException, IOException {

        addImage(new File(filename), originalSize, name, location, overwrite);
    }

    public void save() throws FileNotFoundException, IOException {
        logger.log(Level.INFO, "Saving workbook to '" + excelFile.getCanonicalPath() + "'");
        FileOutputStream fos = new FileOutputStream(excelFile);
        workbook.write(fos);
        fos.close();
    }

    /**
     * UTILITY FUNCTIONS
     */

    private Name getName(String name) {
        Name cname = workbook.getName(name);
        if(cname != null)
            return cname;
        else
            logger.log(Level.SEVERE, "Name '" + name + "' does not exist!");
            throw new IllegalArgumentException("Name '" + name + "' does not exist!");
    }

    private void defineOrOverwriteName(String name, String location, boolean overwrite) {
        AreaReference areaReference = new AreaReference(location);
        String sheetName = areaReference.getFirstCell().getSheetName();

        if(isNameExisting(name)) {
            logger.log(Level.INFO, "Name already exists");
            if(overwrite) {
                // Name already exists but we overwrite --> remove
                logger.log(Level.INFO, "Specified to overwrite named region if already existing, " +
                        "therefore remove existing name");
                removeName(name);
            } else {
                // Name already exists but we don't want to overwrite --> error
                logger.log(Level.SEVERE, "Specified named region already exists - specified to not overwrite");
                throw new IllegalArgumentException("Specified named region '" + name + "' already exists!");
            }
        }

        createSheet(sheetName);
        createName(name, location);
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

    /*
    private Cell getCell(CellReference cref) {
        Sheet sheet = workbook.getSheet(cref.getSheetName());
        return getCell(sheet, cref.getRow(), cref.getCol());
    }
     * 
     */

    private void setMissing(Cell cell) {
        cell.setCellType(Cell.CELL_TYPE_BLANK);
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
                            throw new IllegalArgumentException("Unknown column type detected!");
                    }
                }
                break;
            case PREDEFINED:
                // In case of a header, determine header styles
                if(data.hasColumnHeader()) {
                    for(int i = 0; i < data.columns(); i++) {
                        cstyles.put(HEADER + i, getCell(sheet, startRow, startCol + i).getCellStyle());
                    }
                }
                int styleRow = startRow + (data.hasColumnHeader() ? 1 : 0);
                for(int i = 0; i < data.columns(); i++) {
                    Cell cell = getCell(sheet, styleRow, startCol + i);
                    cstyles.put(COLUMN + i, cell.getCellStyle());
                }
                break;
            case STYLE_NAME:
                if(isHSSF()) {
                    HSSFWorkbook wb = (HSSFWorkbook) workbook;
                    if(data.hasColumnHeader()) {
                        for(int i = 0; i < data.columns(); i++)
                            cstyles.put(HEADER + i, WorkbookUtils.findCellStyleByName(wb, styleName + "_" + HEADER_STYLE));
                    }
                    for(int i = 0; i < data.columns(); i++) {
                        switch(data.getColumnType(i)) {
                            case Boolean:
                                cstyles.put(COLUMN + i, WorkbookUtils.findCellStyleByName(wb, styleName + "_" + BOOLEAN_STYLE));
                                break;
                            case DateTime:
                                cstyles.put(COLUMN + i, WorkbookUtils.findCellStyleByName(wb, styleName + "_" + DATETIME_STYLE));
                                break;
                            case Numeric:
                                cstyles.put(COLUMN + i, WorkbookUtils.findCellStyleByName(wb, styleName + "_" + NUMERIC_STYLE));
                                break;
                            case String:
                                cstyles.put(COLUMN + i, WorkbookUtils.findCellStyleByName(wb, styleName + "_" + STRING_STYLE));
                                break;
                            default:
                                throw new IllegalArgumentException("Unknown column type detected!");
                        }
                    }
                } else
                    throw new IllegalArgumentException("Style name action not supported for XSSF worksheets!");
                break;
            default:
                throw new IllegalArgumentException("Style action not supported!");
        }

        return cstyles;
    }

    /***
     * FACTORY METHODS
     */

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
    public static Workbook getWorkbook(File excelFile) throws FileNotFoundException, IOException, InvalidFormatException {
        Workbook wb;

        if(excelFile.exists()) {
            logger.log(Level.INFO, "Creating XLConnect workbook instance for existing file '" + excelFile.getCanonicalPath() + "'");
            wb = new Workbook(excelFile);
        } else {
            logger.log(Level.INFO, "Creating XLConnect workbook instance for new file '" + excelFile.getCanonicalPath() + "'");
            String filename = excelFile.getName().toLowerCase();
            if(filename.endsWith(".xls")) {
                wb = new Workbook(excelFile, SpreadsheetVersion.EXCEL97);
            } else if(filename.endsWith(".xlsx")) {
                wb = new Workbook(excelFile, SpreadsheetVersion.EXCEL2007);
            } else
                throw new IllegalArgumentException("File extension not supported! Only *.xls and *.xlsx are allowd!");
        }

        logger.log(Level.INFO, "Excel version: " + (wb.isHSSF() ? SpreadsheetVersion.EXCEL97.toString() : SpreadsheetVersion.EXCEL2007.toString()));
        return wb;
    }

    public static Workbook getWorkbook(String filename) throws FileNotFoundException, IOException, InvalidFormatException {
        return Workbook.getWorkbook(new File(filename));
    }
}
