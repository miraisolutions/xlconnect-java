/*
 *
    XLConnect
    Copyright (C) 2010-2024 Mirai Solutions GmbH

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

import com.miraisolutions.xlconnect.data.*;
import com.miraisolutions.xlconnect.utils.DateTimeFormatter;
import com.miraisolutions.xlconnect.utils.RPOSIXDateTimeFormatter;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.file.Files;
import java.util.*;
import java.util.function.Consumer;
import java.util.stream.IntStream;
import java.util.stream.Collectors;


/**
 * Class representing a Microsoft Excel Workbook for XLConnect
 */
public final class Workbook {

    // Prefix
    private final static String HEADER = "Header";
    private final static String COLUMN = "Column";
    private final static String SEP = ".";

    private final static String HEADER_STYLE = "XLConnect.Header";
    private final static String NUMERIC_STYLE = "XLConnect.Numeric";
    private final static String STRING_STYLE = "XLConnect.String";
    private final static String BOOLEAN_STYLE = "XLConnect.Boolean";
    private final static String DATETIME_STYLE = "XLConnect.DateTime";

    // Formatter
    // NOTE: currently fixed to a RPOSIXDateTimeFormatter
    public final static DateTimeFormatter dateTimeFormatter = new RPOSIXDateTimeFormatter();

    static {
        ZipSecureFile.setMinInflateRatio(0.001);
    }

    // Apache POI workbook instance
    private final org.apache.poi.ss.usermodel.Workbook workbook;
    // Underlying file instance
    private File excelFile;
    // Style action
    private StyleAction styleAction = StyleAction.XLCONNECT;
    // Style name prefix
    private String styleNamePrefix = null;
    /* Missing value strings;
       first element is used as missing value string when writing data
       (null means blank/empty cell)
     */
    private Object[] missingValue = new Object[]{null};
    // Default cell styles
    private final Map<String, CellStyle> defaultStyles =
            new HashMap<>(5);
    // Styles per data type
    private final Map<DataType, CellStyle> dataTypeStyles = new EnumMap<>(DataType.class);

    // Data format map
    private final Map<DataType, String> dataFormatMap = new EnumMap<>(DataType.class);


    // Behavior when detecting an error cell
    // WARN means returning a missing value and registering a warning
    private ErrorBehavior onErrorCell = ErrorBehavior.WARN;

    // This is used to support the warnings mechanism on the R side
    private ArrayList<String> warnings = new ArrayList<>();

    private Workbook(File excelFile, String password) throws IOException {
        /*
         * NOTE: We are using a FileInputStream since otherwise using 'save' multiple times would cause
         * a JVM crash as described here: https://bz.apache.org/bugzilla/show_bug.cgi?id=53515
         */
        this.workbook = WorkbookFactory.create(Files.newInputStream(excelFile.toPath()), password);
        this.excelFile = excelFile;
        init();
    }

    private Workbook(File excelFile) throws IOException {
        /*
         * NOTE: We are using a FileInputStream since otherwise using 'save' multiple times would cause
         * a JVM crash as described here: https://bz.apache.org/bugzilla/show_bug.cgi?id=53515
         */
        this.workbook = WorkbookFactory.create(Files.newInputStream(excelFile.toPath()));
        this.excelFile = excelFile;
        init();
    }

    private Workbook(File excelFile, SpreadsheetVersion version) {
        switch (version) {
            case EXCEL97:
                this.workbook = new HSSFWorkbook();
                break;
            case EXCEL2007:
                this.workbook = new XSSFWorkbook();
                break;
            default:
                throw new IllegalArgumentException("Spreadsheet version not supported!");
        }

        this.excelFile = excelFile;
        init();
    }

    private void init() {
        initDefaultDataFormats();
        initDefaultStyles();
    }

    private void initDefaultDataFormats() {
        dataFormatMap.put(DataType.Boolean, "General");
        dataFormatMap.put(DataType.DateTime, "mm/dd/yyyy hh:mm:ss");
        dataFormatMap.put(DataType.Numeric, "General");
        dataFormatMap.put(DataType.String, "General");
    }

    private CellStyle initGeneralStyle(String name) {
        CellStyle style = getCellStyle(name);
        if (style == null) {
            style = createCellStyle(name);
            style.setDataFormat("General");
            style.setWrapText(true);
        }
        return style;
    }

    private void initDefaultStyles() {
        // Header style
        CellStyle headerStyle = getCellStyle(HEADER_STYLE);
        if (headerStyle == null) {
            headerStyle = initGeneralStyle(HEADER_STYLE);
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        }

        // String / boolean / numeric style
        CellStyle stringStyle = initGeneralStyle(STRING_STYLE);
        dataTypeStyles.put(DataType.String, stringStyle);
        CellStyle numericStyle = initGeneralStyle(NUMERIC_STYLE);
        dataTypeStyles.put(DataType.Numeric, numericStyle);
        CellStyle booleanStyle = initGeneralStyle(BOOLEAN_STYLE);
        dataTypeStyles.put(DataType.Boolean, booleanStyle);

        // Date style
        CellStyle dateStyle = getCellStyle(DATETIME_STYLE);
        if (dateStyle == null) {
            dateStyle = createCellStyle(DATETIME_STYLE);
            dateStyle.setDataFormat(dataFormatMap.get(DataType.DateTime));
            dateStyle.setWrapText(true);
        }
        dataTypeStyles.put(DataType.DateTime, dateStyle);

        defaultStyles.put(HEADER_STYLE, headerStyle);
        defaultStyles.put(STRING_STYLE, stringStyle);
        defaultStyles.put(NUMERIC_STYLE, numericStyle);
        defaultStyles.put(BOOLEAN_STYLE, booleanStyle);
        defaultStyles.put(DATETIME_STYLE, dateStyle);
    }

    public void setCellStyleForDataType(DataType type, CellStyle cs) {
        dataTypeStyles.put(type, cs);
    }

    public CellStyle getCellStyleForDataType(DataType type) {
        return dataTypeStyles.get(type);
    }

    public void setDataFormat(DataType type, String format) {
        dataFormatMap.put(type, format);
    }

    public void setStyleAction(StyleAction styleAction) {
        this.styleAction = styleAction;
    }

    public void setStyleNamePrefix(String styleNamePrefix) {
        this.styleNamePrefix = styleNamePrefix;
    }

    public String[] getSheets() {
        return IntStream.range(0, workbook.getNumberOfSheets())
                .mapToObj(workbook::getSheetName)
                .toArray(String[]::new);
    }

    public int getSheetPos(String sheetName) {
        return workbook.getSheetIndex(sheetName);
    }

    public void setSheetPos(String sheetName, int pos) {
        workbook.setSheetOrder(sheetName, pos);
    }

    public String[] getDefinedNames(boolean validOnly, String worksheetScope) {

        return workbook.getAllNames().stream()
                .filter(n -> (!validOnly || isValidNamedRegion(n)) &&
                        (worksheetScope == null || n.getSheetIndex() == getSheetIndexForScope(worksheetScope)))
                .map(Name::getNameName).toArray(String[]::new);
    }


    private boolean isValidNamedRegion(Name region) {
        return !region.isDeleted() && hasValidWorkSheet(region);
    }

    /**
     * Returns the sheet index for the worksheet name or -1 if the sheet name is ""
     * @param worksheetScope the worksheet name
     * @throws NoSuchElementException if no worksheet exists with this name
     */
    private int getSheetIndexForScope(String worksheetScope) {
        if (worksheetScope.isEmpty()) return -1;
        int index = workbook.getSheetIndex(worksheetScope);
        if (index < 0) throw new NoSuchElementException("Worksheet " + worksheetScope + " was not found!");
        else return index;
    }

    private boolean hasValidWorkSheet(Name region) {
        String sheetName = null;
        try {
            sheetName = region.getSheetName();
        } catch (Exception ignored) {
        }
        return (sheetName != null && !sheetName.isEmpty());
    }

    public boolean existsSheet(String name) {
        return workbook.getSheet(name) != null;
    }

    public boolean existsName(String name, String worksheetScope) {
        try {
            getName(name, worksheetScope);
            return true;
        }
        catch (IllegalArgumentException ignored) {
            return false;
        }
    }


    public void createSheet(String name) {
        if (name.length() > 31)
            throw new IllegalArgumentException("Sheet names are not allowed to contain more than 31 characters!");

        if (workbook.getSheetIndex(name) < 0)
            workbook.createSheet(name);
    }

    public void removeSheet(int sheetIndex) {
        if (sheetIndex > -1 && sheetIndex < workbook.getNumberOfSheets()) {
            setAlternativeActiveSheet(sheetIndex);
            workbook.removeSheetAt(sheetIndex);
        }
    }

    public void removeSheet(String name) {
        removeSheet(workbook.getSheetIndex(name));
    }

    public void renameSheet(int sheetIndex, String newName) {
        renameSheet(workbook.getSheetName(sheetIndex), newName);
    }

    public void renameSheet(String name, String newName) {
        workbook.setSheetName(workbook.getSheetIndex(name), newName);
    }

    public void cloneSheet(int index, String newName) {
        Sheet sheet = workbook.cloneSheet(index);
        String originalName = workbook.getSheetName(index);
        workbook.setSheetName(workbook.getSheetIndex(sheet), newName);
        // Copy names (named ranges) that are scoped to the original sheet, adapting the scope.
        List<Name> originalNamedRanges = workbook.getAllNames().stream().filter(name -> name.getSheetIndex() == index).collect(Collectors.toList());
        // we have to collect the original names *then* add the new names, otherwise we get concurrency issues creating names while iterating over them
        originalNamedRanges.forEach(namedRange ->
                createName(namedRange.getNameName(), newName, namedRange.getRefersToFormula(), false));
    }

    public void cloneSheet(String name, String newName) {
        cloneSheet(workbook.getSheetIndex(name), newName);
    }

    public void createName(String name,  String worksheetScope, String formula, boolean overwrite) {
        if (existsName(name, worksheetScope)) {
            if (overwrite) {
                // Name already exists but we overwrite --> remove
                removeName(name, worksheetScope);
            } else {
                // Name already exists, but we don't want to overwrite --> error
                throw new IllegalArgumentException("Specified name '" + name + "' already exists in " + worksheetScope);
            }
        }

        Name cname = workbook.createName();
        if(worksheetScope != null) {
            int sheetIndex = getSheetIndexForScope(worksheetScope);
            if(sheetIndex >= 0) cname.setSheetIndex(sheetIndex);
        }
        try {
            cname.setNameName(name);
            cname.setRefersToFormula(formula);
        } catch (Exception e) {
            // --> Clean up (= remove) name
            // Need to set dummy name in order to be able to remove it ...
            workbook.removeName(cname);
            throw new IllegalArgumentException(e);
        }
    }

    public void removeName(String name, String worksheetScope) {
        if (existsName(name, worksheetScope)) {
            Name cname = getName(name, worksheetScope);
            workbook.removeName(cname);
        }
    }

    public String getReferenceFormula(String name, String worksheetScope) {
        return getName(name, worksheetScope).getRefersToFormula();
    }

    // Keep for backwards compatibility
    public int[] getReferenceCoordinates(String name) {
        return getReferenceCoordinatesForName(name, null);
    }

    public int[] getReferenceCoordinatesForName(String name, String worksheetScope) {
        Name cname = getName(name, worksheetScope);
        AreaReference aref = new AreaReference(cname.getRefersToFormula(), workbook.getSpreadsheetVersion());
        // Get upper left corner
        CellReference first = aref.getFirstCell();
        // Get lower right corner
        CellReference last = aref.getLastCell();
        int top = first.getRow();
        int bottom = last.getRow();
        int left = first.getCol();
        int right = last.getCol();
        return new int[]{top, left, bottom, right};
    }

    public String[] getTables(int sheetIndex) {
        if (isXSSF()) {
            XSSFSheet s = (XSSFSheet) getSheet(sheetIndex);
            return s.getTables().stream()
                    .map(XSSFTable::getName)
                    .toArray(String[]::new);
        } else {
            return new String[0];
        }
    }

    public String[] getTables(String sheetName) {
        return getTables(workbook.getSheetIndex(sheetName));
    }

    public int[] getReferenceCoordinatesForTable(int sheetIndex, String tableName) {
        if (!isXSSF()) {
            throw new IllegalArgumentException("Tables are not supported with this file format");
        }
        XSSFSheet s = (XSSFSheet) getSheet(sheetIndex);
        for (XSSFTable t : s.getTables()) {
            if (tableName.equals(t.getName())) {
                CellReference start = t.getStartCellReference();
                CellReference end = t.getEndCellReference();
                int top = start.getRow();
                int bottom = end.getRow();
                int left = start.getCol();
                int right = end.getCol();
                return new int[]{top, left, bottom, right};
            }
        }
        throw new IllegalArgumentException("Could not find table '" + tableName + "'!");
    }

    public int[] getReferenceCoordinatesForTable(String sheetName, String tableName) {
        return getReferenceCoordinatesForTable(workbook.getSheetIndex(sheetName), tableName);
    }

    private void writeData(DataFrame data, Sheet sheet, int startRow, int startCol, boolean header, boolean overwriteFormulaCells) {
        // Get styles
        Map<String, CellStyle> styles = getStyles(data, sheet, startRow, startCol);

        Consumer<Cell> maybeClearFormula = overwriteFormulaCells ? (cell ->
        {
            if (cell.getCellType() == CellType.FORMULA) {
                cell.removeFormula();
            }
        }
        ) : (cell -> {
        });
        // Define row & column index variables
        int rowIndex = startRow;
        int colIndex = startCol;

        // In case of column headers ...
        if (header && data.hasColumnHeader()) {
            // For each column write corresponding column name
            for (int i = 0; i < data.columns(); i++) {
                Cell cell = getCell(sheet, rowIndex, colIndex + i);
                cell.setCellValue(data.getColumnName(i));
                setCellStyle(cell, styles.get(HEADER + i));
            }

            ++rowIndex;
        }

        // For each column of data
        for (int i = 0; i < data.columns(); i++) {
            // Get column style
            CellStyle cs = styles.get(COLUMN + i);
            Column col = data.getColumn(i);
            // Depending on column type ...
            switch (data.getColumnType(i)) {
                case Numeric:
                    double[] doubleValues = col.getNumericData();
                    for (int j = 0; j < data.rows(); j++) {
                        Cell cell = getCell(sheet, rowIndex + j, colIndex);
                        maybeClearFormula.accept(cell);
                        if (col.isMissing(j))
                            setMissing(cell);
                        else {
                            if (Double.isInfinite(doubleValues[j])) {
                                cell.setCellErrorValue(FormulaError.NA.getCode());
                            } else {
                                cell.setCellValue(doubleValues[j]);
                            }
                            setCellStyle(cell, cs);
                        }
                    }
                    break;
                case String:
                    String[] stringValues = col.getStringData();
                    for (int j = 0; j < data.rows(); j++) {
                        Cell cell = getCell(sheet, rowIndex + j, colIndex);
                        maybeClearFormula.accept(cell);
                        if (col.isMissing(j))
                            setMissing(cell);
                        else {
                            cell.setCellValue(stringValues[j]);
                            setCellStyle(cell, cs);
                        }
                    }
                    break;
                case Boolean:
                    boolean[] booleanValues = col.getBooleanData();
                    for (int j = 0; j < data.rows(); j++) {
                        Cell cell = getCell(sheet, rowIndex + j, colIndex);
                        maybeClearFormula.accept(cell);
                        if (col.isMissing(j))
                            setMissing(cell);
                        else {
                            cell.setCellValue(booleanValues[j]);
                            setCellStyle(cell, cs);
                        }
                    }
                    break;
                case DateTime:
                    Date[] dateValues = col.getDateTimeData();
                    for (int j = 0; j < data.rows(); j++) {
                        Cell cell = getCell(sheet, rowIndex + j, colIndex);
                        maybeClearFormula.accept(cell);
                        if (col.isMissing(j))
                            setMissing(cell);
                        else {
                            cell.setCellValue(dateValues[j]);
                            setCellStyle(cell, cs);
                        }
                    }
                    break;
                default:
                    throw new IllegalArgumentException("Unknown column type detected!");
            }

            ++colIndex;
        }
    }


    private DataFrame readData(Sheet sheet, int startRow, int startCol, int nrows, int ncols, boolean header,
                               ReadStrategy readStrategy, DataType[] colTypes, boolean forceConversion, String dateTimeFormat,
                               boolean takeCached, int[] subset) {

        DataFrame data = new DataFrame();
        int[] colset;

        // Formula evaluator - only if we don't want to take cached values
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        if (!takeCached) evaluator.clearAllCachedResultValues();

        if (subset == null) {
            colset = IntStream.range(0, ncols).toArray();
        } else {
            colset = subset;
        }

        ColumnBuilder cb;
        switch (readStrategy) {
            case DEFAULT:
                cb = new DefaultColumnBuilder(nrows, forceConversion, takeCached, evaluator, onErrorCell,
                        missingValue, dateTimeFormat);
                break;
            case FAST:
                cb = new FastColumnBuilder(nrows, forceConversion, takeCached, evaluator, onErrorCell,
                        dateTimeFormat);
                break;
            default:
                throw new IllegalArgumentException("Unknown read strategy!");
        }

        // Determine header column
        String[] columnHeaders = new String[colset.length];
        if (header) {
            ColumnBuilder cbHeader = new DefaultColumnBuilder(1, true, takeCached, evaluator, onErrorCell, missingValue, dateTimeFormat);
            for (int col : colset) {
                cbHeader.addCell(getCell(sheet, startRow, startCol + col, false));
            }
            columnHeaders = cbHeader.buildStringColumn().getStringData();
        }
        // Replace missing column headers
        for (int i = 0; i < columnHeaders.length; i++) {
            if (columnHeaders[i] == null) {
                columnHeaders[i] = "Col" + (colset[i] + 1);
            }
        }

        // Loop over columns
        for (int i = 0; i < colset.length; i++) {
            int col = colset[i];
            int colIndex = startCol + col;
            String columnHeader = columnHeaders[i];

            // Prepare column builder for new set of rows
            cb.clear();

            // Loop over rows
            Row r;
            for (int row = header ? 1 : 0; row < nrows; row++) {
                int rowIndex = startRow + row;

                // Cell cell = getCell(sheet, rowIndex, colIndex, false);
                Cell cell = ((r = sheet.getRow(rowIndex)) == null) ? null : r.getCell(colIndex);
                cb.addCell(cell);
            }

            DataType columnType = ((colTypes != null) && (colTypes.length > 0)) ? colTypes[col % colTypes.length] :
                    cb.determineColumnType();
            switch (columnType) {
                case Boolean:
                    data.addColumn(columnHeader, cb.buildBooleanColumn());
                    break;
                case DateTime:
                    data.addColumn(columnHeader, cb.buildDateTimeColumn());
                    break;
                case Numeric:
                    data.addColumn(columnHeader, cb.buildNumericColumn());
                    break;
                case String:
                    data.addColumn(columnHeader, cb.buildStringColumn());
                    break;
                default:
                    throw new IllegalArgumentException("Unknown data type detected!");

            }
            // Collect column builder warnings
            this.warnings.addAll(cb.getWarnings());
        }

        return data;
    }


    public void onErrorCell(ErrorBehavior eb) {
        this.onErrorCell = eb;
    }

    public void writeNamedRegion(DataFrame data, String name, boolean header, boolean overwriteFormulaCells, String worksheetScope) {
        Name cname = getName(name, worksheetScope);
        checkName(cname);

        // Get sheet where name is defined in
        Sheet sheet = workbook.getSheet(cname.getSheetName());

        AreaReference aref = new AreaReference(cname.getRefersToFormula(), workbook.getSpreadsheetVersion());
        // Get upper left corner
        CellReference topLeft = aref.getFirstCell();

        // Compute bottom right cell coordinates
        int bottomRightRow = Math.max(topLeft.getRow() + data.rows() - 1, topLeft.getRow());
        if (header && data.rows() > 0) ++bottomRightRow;
        int bottomRightCol = Math.max(topLeft.getCol() + data.columns() - 1, topLeft.getCol());
        // Create bottom right cell reference
        CellReference bottomRight = new CellReference(sheet.getSheetName(), bottomRightRow,
                bottomRightCol, true, true);

        // Define named range area
        aref = new AreaReference(topLeft, bottomRight, workbook.getSpreadsheetVersion());
        // Redefine named range
        cname.setRefersToFormula(aref.formatAsString());

        writeData(data, sheet, topLeft.getRow(), topLeft.getCol(), header, overwriteFormulaCells);
    }

    public DataFrame readNamedRegion(String name, String worksheetScope, boolean header, ReadStrategy readStrategy, DataType[] colTypes,
                                     boolean forceConversion, String dateTimeFormat, boolean takeCached, int[] subset) {
        Name cname = getName(name, worksheetScope);
        checkName(cname);

        // Get sheet where name is defined in
        Sheet sheet = workbook.getSheet(cname.getSheetName());

        AreaReference aref = new AreaReference(cname.getRefersToFormula(), workbook.getSpreadsheetVersion());
        // Get name corners (top left, bottom right)
        CellReference topLeft = aref.getFirstCell();
        CellReference bottomRight = aref.getLastCell();

        // Determine number of rows and columns
        int nrows = bottomRight.getRow() - topLeft.getRow() + 1;
        int ncols = bottomRight.getCol() - topLeft.getCol() + 1;

        return readData(sheet, topLeft.getRow(), topLeft.getCol(), nrows, ncols, header, readStrategy, colTypes,
                forceConversion, dateTimeFormat, takeCached, subset);
    }

    public DataFrame readTable(int worksheetIndex, String tableName, boolean header, ReadStrategy readStrategy,
                               DataType[] colTypes, boolean forceConversion, String dateTimeFormat, boolean takeCached, int[] subset) {
        if (!isXSSF()) throw new IllegalArgumentException("Tables are not supported with this file format!");
        XSSFSheet s = (XSSFSheet) getSheet(worksheetIndex);
        int[] coords = getReferenceCoordinatesForTable(worksheetIndex, tableName);
        int nrows = coords[2] - coords[0] + 1;
        int ncols = coords[3] - coords[1] + 1;
        return readData(s, coords[0], coords[1], nrows, ncols, header, readStrategy, colTypes, forceConversion, dateTimeFormat,
                takeCached, subset);
    }

    public DataFrame readTable(String worksheetName, String tableName, boolean header, ReadStrategy readStrategy,
                               DataType[] colTypes, boolean forceConversion, String dateTimeFormat, boolean takeCached, int[] subset) {
        return readTable(workbook.getSheetIndex(worksheetName), tableName, header, readStrategy, colTypes,
                forceConversion, dateTimeFormat, takeCached, subset);
    }

    /**
     * Writes a data frame into the specified worksheet index at the specified location
     *
     * @param data           Data frame to be written to the worksheet
     * @param worksheetIndex Worksheet index (0-based)
     * @param startRow       Start row (row index of top left cell)
     * @param startCol       Start column (column index of top left cell)
     * @param header         If true, column headers are written, otherwise not
     */
    public void writeWorksheet(DataFrame data, int worksheetIndex, int startRow, int startCol, boolean header, boolean overwriteFormulaCells) {
        Sheet sheet = workbook.getSheetAt(worksheetIndex);
        writeData(data, sheet, startRow, startCol, header, overwriteFormulaCells);
    }

    public void writeWorksheet(DataFrame data, String worksheetName, int startRow, int startCol, boolean header, boolean overwriteFormulaCells) {
        int sheetIndex = workbook.getSheetIndex(worksheetName);
        if(sheetIndex < 0)
            throw new NoSuchElementException("Worksheet " + worksheetName + " was not found!");
        writeWorksheet(data, sheetIndex, startRow, startCol, header, overwriteFormulaCells);
    }

    public void writeWorksheet(DataFrame data, int worksheetIndex, boolean header, boolean overwriteFormulaCells) {
        writeWorksheet(data, worksheetIndex, 0, 0, header, overwriteFormulaCells);
    }

    public void writeWorksheet(DataFrame data, String worksheetName, boolean header, boolean overwriteFormulaCells) {
        writeWorksheet(data, worksheetName, 0, 0, header, overwriteFormulaCells);
    }

    /**
     * Reads data from a worksheet. Data regions can be narrowed down by specifying corresponding row and column ranges.
     * Limits specified as negative integers will be automatically determined. The rules for automatically determining
     * the ranges are as follows:
     * <p>
     * - If start row < 0: get first row on sheet
     * - If end row < 0: get last row on sheet
     * - If start column < 0: get column of first (non-null) cell in start row
     * - If end column < 0: get max column between start row and end row
     *
     * @param worksheetIndex  Worksheet index
     * @param startRow        Start row
     * @param startCol        Start column
     * @param endRow          End row
     * @param endCol          End column
     * @param header          If true, assume header, otherwise not
     * @param colTypes        Column data types
     * @param forceConversion Should conversion to a less generic data type be forced?
     * @param dateTimeFormat  Date/time format used when converting between Date and String
     * @return Data Frame
     */
    public DataFrame readWorksheet(int worksheetIndex, int startRow, int startCol, int endRow, int endCol, boolean header,
                                   ReadStrategy readStrategy, DataType[] colTypes, boolean forceConversion, String dateTimeFormat,
                                   boolean takeCached, int[] subset, boolean autofitRow, boolean autofitCol) {
        Sheet sheet = workbook.getSheetAt(worksheetIndex);
        int[] boundingBox = getBoundingBox(worksheetIndex, startRow, startCol, endRow, endCol, autofitRow, autofitCol);
        startRow = boundingBox[0];
        startCol = boundingBox[1];
        endRow = boundingBox[2];
        endCol = boundingBox[3];

        int nrows = startRow < 0 ? 0 : (endRow - startRow) + 1;
        int ncols = startCol < 0 ? 0 : (endCol - startCol) + 1;
        if (nrows == 0 || ncols == 0) {
            this.warnings.add("Data frame contains " + nrows + " rows and " + ncols + " columns!");
        }

        return readData(sheet, startRow, startCol, nrows, ncols, header, readStrategy, colTypes, forceConversion, dateTimeFormat,
                takeCached, subset);
    }

    public DataFrame readWorksheet(String worksheetName, int startRow, int startCol, int endRow, int endCol, boolean header,
                                   ReadStrategy readStrategy, DataType[] colTypes, boolean forceConversion, String dateTimeFormat, boolean takeCached,
                                   int[] subset, boolean autofitRow, boolean autofitCol) {
        return readWorksheet(workbook.getSheetIndex(worksheetName), startRow, startCol, endRow, endCol, header, readStrategy,
                colTypes, forceConversion, dateTimeFormat, takeCached, subset, autofitRow, autofitCol);
    }

    public DataFrame readWorksheet(String worksheetName, int startRow, int startCol, int endRow, int endCol, boolean header) {
        return readWorksheet(worksheetName, startRow, startCol, endRow, endCol, header, ReadStrategy.DEFAULT, 
                null, false, "", false, null, true, true);
    }
    
    public DataFrame readWorksheet(String worksheetName, int startRow, int startCol, int endRow, int endCol, boolean header,
            boolean autofitRow, boolean autofitCol) {
        return readWorksheet(worksheetName, startRow, startCol, endRow, endCol, header, ReadStrategy.DEFAULT,
                null, false, "", false, null, autofitRow, autofitCol);
    }

    public DataFrame readWorksheet(String worksheetName, boolean header, ReadStrategy readStrategy, DataType[] colTypes, 
            boolean forceConversion, String dateTimeFormat) {
        return readWorksheet(worksheetName, -1, -1, -1, -1, header, readStrategy, colTypes, forceConversion, 
                dateTimeFormat, false, null, true, true);
    }

    public DataFrame readWorksheet(String worksheetName, boolean header) {
        return readWorksheet(worksheetName, header, ReadStrategy.DEFAULT, null, false, "");
    }

    public void addImage(File imageFile, String name, String worksheetScope, boolean originalSize) throws IOException {
        Name cname = getName(name, worksheetScope);

        // Get sheet where name is defined in
        Sheet sheet = workbook.getSheet(cname.getSheetName());

        AreaReference aref = new AreaReference(cname.getRefersToFormula(), workbook.getSpreadsheetVersion());
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
        } else
            throw new IllegalArgumentException("Image type \""+ filename.substring(filename.lastIndexOf('.')+1) +"\" not supported!");
        InputStream is = Files.newInputStream(imageFile.toPath());
        byte[] bytes = IOUtils.toByteArray(is);
        int imageIndex = workbook.addPicture(bytes, imageType);
        is.close();

        Drawing<?> drawing;
        if (isHSSF()) {
            drawing = ((HSSFSheet) sheet).getDrawingPatriarch();
            if (drawing == null) {
                drawing = sheet.createDrawingPatriarch();
            }
        } else if (isXSSF()) {
            drawing = ((XSSFSheet) sheet).createDrawingPatriarch();
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
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);

        Picture picture = drawing.createPicture(anchor, imageIndex);
        if(originalSize) picture.resize();
    }

    public void addImage(String filename, String name, String worksheetScope, boolean originalSize) throws IOException {
        addImage(new File(filename), name, worksheetScope, originalSize);
    }

    public CellStyle createCellStyle(String name) {
        if (getCellStyle(name) == null) {
            if (isHSSF()) {
                return HCellStyle.create((HSSFWorkbook) workbook, name);
            } else if (isXSSF()) {
                return XCellStyle.create((XSSFWorkbook) workbook, name);
            } else {
                throw new RuntimeException("Unsupported workbook format.");
            }
        } else
            throw new IllegalArgumentException("Cell style with name '" + name + "' already exists!");
    }

    public CellStyle createCellStyle() {
        return createCellStyle(null);
    }

    public int getActiveSheetIndex() {
        if (workbook.getNumberOfSheets() < 1)
            return -1;
        else
            return workbook.getActiveSheetIndex();
    }

    public String getActiveSheetName() {
        if (workbook.getNumberOfSheets() < 1)
            return null;
        else
            return workbook.getSheetName(workbook.getActiveSheetIndex());
    }

    public void setActiveSheet(int sheetIndex) {
        workbook.setActiveSheet(sheetIndex);
    }

    public void setActiveSheet(String sheetName) {
        int sheetIndex = workbook.getSheetIndex(sheetName);
        setActiveSheet(sheetIndex);
    }

    public void hideSheet(int sheetIndex, boolean veryHidden) {
        setAlternativeActiveSheet(sheetIndex);
        workbook.setSheetVisibility(sheetIndex, veryHidden ? SheetVisibility.VERY_HIDDEN : SheetVisibility.HIDDEN);
    }

    public void hideSheet(String sheetName, boolean veryHidden) {
        hideSheet(workbook.getSheetIndex(sheetName), veryHidden);
    }

    public void unhideSheet(int sheetIndex) {
        workbook.setSheetVisibility(sheetIndex, SheetVisibility.VISIBLE);
    }

    public void unhideSheet(String sheetName) {
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
        Sheet sheet = getSheet(sheetIndex);
        if (width >= 0)
            sheet.setColumnWidth(columnIndex, width);
        else if (width == -1)
            sheet.autoSizeColumn(columnIndex);
        else
            sheet.setColumnWidth(columnIndex, sheet.getDefaultColumnWidth() * 256);
    }

    public void setColumnWidth(String sheetName, int columnIndex, int width) {
        setColumnWidth(workbook.getSheetIndex(sheetName), columnIndex, width);
    }

    public void setRowHeight(int sheetIndex, int rowIndex, float height) {
        Sheet sheet = getSheet(sheetIndex);
        Row r = sheet.getRow(rowIndex);
        if (r == null)
            r = getSheet(sheetIndex).createRow(rowIndex);

        if (height >= 0)
            r.setHeightInPoints(height);
        else
            r.setHeightInPoints(sheet.getDefaultRowHeightInPoints());
    }

    public void setRowHeight(String sheetName, int rowIndex, float height) {
        setRowHeight(workbook.getSheetIndex(sheetName), rowIndex, height);
    }

    public void save(OutputStream os) throws IOException {
        workbook.write(os);
    }

    public void save(File f) throws IOException {
        this.excelFile = f;
        FileOutputStream fos = new FileOutputStream(f, false);
        save(fos);
        fos.close();
    }

    public void save(String file) throws IOException {
        save(new File(file));
    }

    public void save() throws IOException {
        save(excelFile);
    }

    Name getName(String name) {
        Name cname = workbook.getName(name);
        if (cname != null)
            return cname;
        else
            throw new IllegalArgumentException("Name '" + name + "' does not exist!");
    }

    List<Name> getNames(String name) {
        return Collections.unmodifiableList(workbook.getNames(name));
    }

    private Name getName(String name, String worksheetScope) {
        if (worksheetScope == null)
            return getName(name);
        int sheetIndex = getSheetIndexForScope(worksheetScope);
        List<Name> cNames = getNames(name);
        for (Name n : cNames)
            if (n.getSheetIndex() == sheetIndex)
                return n;

        StringBuffer names = new StringBuffer();
        String worksheetScopeDisplay = worksheetScope.isEmpty() ? "global scope" : worksheetScope;
        cNames.forEach(n -> names.append(n.getSheetIndex() >= 0 ? workbook.getSheetName(n.getSheetIndex()) : "global scope").append(";"));
        throw new IllegalArgumentException("Name '" + name + "' was not specified in worksheet '" + worksheetScopeDisplay + "'! " +
                "Found in sheets: " + names);
    }

    // Checks only if the reference as such is valid
    private boolean isValidReference(String reference) {
        return reference != null && !reference.startsWith("#REF!") && !reference.startsWith("#NULL!");
    }

    private void checkName(Name name) {
        if (!isValidReference(name.getRefersToFormula()))
            throw new IllegalArgumentException("Name '" + name.getNameName() + "' has invalid reference!");
        else if (!existsSheet(name.getSheetName())) {
            // The reference as such is valid, but it doesn't point to a (existing) sheet ...
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
        if (row == null) {
            if (create) {
                row = sheet.createRow(rowIndex);
            } else return null;
        }
        // Get or create cell
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            if (create) {
                cell = row.createCell(colIndex);
            } else return null;
        }

        return cell;
    }

    private Cell getCell(Sheet sheet, int rowIndex, int colIndex) {
        return getCell(sheet, rowIndex, colIndex, true);
    }

    private Sheet getSheet(int sheetIndex) {
        if (sheetIndex < 0 || sheetIndex >= workbook.getNumberOfSheets())
            throw new IllegalArgumentException("Sheet with index " + sheetIndex + " does not exist!");
        return workbook.getSheetAt(sheetIndex);
    }

    private Sheet getSheet(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null)
            throw new IllegalArgumentException("Sheet with name '" + sheetName + "' does not exist!");
        return sheet;
    }

    public void setMissingValue(Object[] values) {
        missingValue = values;
    }

    private void setMissing(Cell cell) {
        if (missingValue.length < 1 || missingValue[0] == null)
            cell.setBlank();
        else {
            if (missingValue[0] instanceof String) {
                cell.setCellValue((String) missingValue[0]);
            } else if (missingValue[0] instanceof Double) {
                cell.setCellValue((Double) missingValue[0]);
            } else {
                cell.setBlank();
                return;
            }

            setCellStyle(cell, DataFormatOnlyCellStyle.get(DataType.String));
        }
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
        if (sheetIndex == getActiveSheetIndex()) {
            // Set active sheet to be first non-hidden/non-very-hidden sheet
            // in the workbook; if there are no such sheets left,
            // then throw an exception
            boolean ok = false;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (i != sheetIndex && !workbook.isSheetHidden(i) && !workbook.isSheetVeryHidden(i)) {
                    setActiveSheet(i);
                    ok = true;
                    break;
                }
            }

            if (!ok) throw new IllegalArgumentException("Cannot hide or remove sheet as there would be no " +
                    "alternative active sheet left!");
        }
    }

    /**
     * Gets a cell style by name.
     *
     * @param name Cell style name
     * @return The corresponding cell style if there exists one with the specified name;
     * null otherwise
     */
    public CellStyle getCellStyle(String name) {
        if (isHSSF()) {
            return HCellStyle.get((HSSFWorkbook) workbook, name);
        } else if (isXSSF()) {
            return XCellStyle.get((XSSFWorkbook) workbook, name);
        }
        return null;
    }

    private CellStyle getCellStyle(Cell cell) {
        return new SSCellStyle(workbook, cell.getCellStyle());
    }

    public boolean existsCellStyle(String name) {
        return getCellStyle(name) != null;
    }

    private void setCellStyle(Cell c, CellStyle cs) {
        if (cs != null) {
            if (cs instanceof HCellStyle) {
                HCellStyle.set((HSSFCell) c, (HCellStyle) cs);
            } else if (cs instanceof XCellStyle) {
                XCellStyle.set((XSSFCell) c, (XCellStyle) cs);
            } else if (cs instanceof DataFormatOnlyCellStyle) {
                CellStyle csx = getCellStyle(c);
                csx.setDataFormat(dataFormatMap.get(((DataFormatOnlyCellStyle) cs).getDataType()));
                SSCellStyle.set(c, (SSCellStyle) csx);
            } else {
                SSCellStyle.set(c, (SSCellStyle) cs);
            }
        }
    }

    private interface CellFunction {
        void apply(Cell cell);
    }

    private void foreachReferencedCell(String formula, CellFunction function) {
        AreaReference aref = new AreaReference(formula, workbook.getSpreadsheetVersion());
        String sheetName = aref.getFirstCell().getSheetName();
        if (sheetName == null) {
            throw new IllegalArgumentException("Invalid formula reference - should be of the form Sheet!A1:B10");
        }
        Sheet sheet = getSheet(sheetName);

        CellReference[] crefs = aref.getAllReferencedCells();
        for (CellReference cref : crefs) {
            Cell cell = getCell(sheet, cref.getRow(), cref.getCol());
            function.apply(cell);
        }
    }

    public void setCellStyle(String formula, final CellStyle cs) {
        foreachReferencedCell(formula, cell -> setCellStyle(cell, cs));
    }

    public void setCellStyle(int sheetIndex, int row, int col, CellStyle cs) {
        Cell c = getCell(getSheet(sheetIndex), row, col);
        setCellStyle(c, cs);
    }

    public void setCellStyle(String sheetName, int row, int col, CellStyle cs) {
        Cell c = getCell(getSheet(sheetName), row, col);
        setCellStyle(c, cs);
    }

    private void setHyperlink(Cell cell, HyperlinkType type, String address) {
        Hyperlink link = workbook.getCreationHelper().createHyperlink(type);
        link.setAddress(address);
        cell.setHyperlink(link);
    }

    public void setHyperlink(String formula, final HyperlinkType type, final String address) {
        foreachReferencedCell(formula, cell -> setHyperlink(cell, type, address));
    }

    public void setHyperlink(int sheetIndex, int row, int col, HyperlinkType type, String address) {
        Cell cell = getCell(getSheet(sheetIndex), row, col);
        setHyperlink(cell, type, address);
    }

    public void setHyperlink(String sheetName, int row, int col, HyperlinkType type, String address) {
        Cell cell = getCell(getSheet(sheetName), row, col);
        setHyperlink(cell, type, address);
    }

    /**
     * Determines the cell styles for headers and columns by column based on the defined style action.
     *
     * @param data     Data frame to be written
     * @param sheet    Worksheet
     * @param startRow Start row in specified sheet for beginning to write the specified data frame
     * @param startCol Start column in specified sheet for beginning to write the specified data frame
     * @return A mapping of header/column indices to cell styles
     */
    private Map<String, CellStyle> getStyles(DataFrame data, Sheet sheet, int startRow, int startCol) {
        Map<String, CellStyle> cstyles = new HashMap<>(data.columns());

        switch (styleAction) {
            case XLCONNECT:
                if (data.hasColumnHeader()) {
                    for (int i = 0; i < data.columns(); i++)
                        cstyles.put(HEADER + i, defaultStyles.get(HEADER_STYLE));
                }
                for (int i = 0; i < data.columns(); i++) {
                    switch (data.getColumnType(i)) {
                        case Boolean:
                            cstyles.put(COLUMN + i, defaultStyles.get(BOOLEAN_STYLE));
                            break;
                        case DateTime:
                            cstyles.put(COLUMN + i, defaultStyles.get(DATETIME_STYLE));
                            break;
                        case Numeric:
                            cstyles.put(COLUMN + i, defaultStyles.get(NUMERIC_STYLE));
                            break;
                        case String:
                            cstyles.put(COLUMN + i, defaultStyles.get(STRING_STYLE));
                            break;
                        default:
                            throw new IllegalArgumentException("Unknown column type detected!");
                    }
                }
                break;
            case DATATYPE:
                if (data.hasColumnHeader()) {
                    for (int i = 0; i < data.columns(); i++)
                        cstyles.put(HEADER + i, defaultStyles.get(HEADER_STYLE));
                }
                for (int i = 0; i < data.columns(); i++) {
                    cstyles.put(COLUMN + i, dataTypeStyles.get(data.getColumnType(i)));
                }
                break;
            case NONE:
                break;
            case PREDEFINED:
                // In case of a header, determine header styles
                if (data.hasColumnHeader()) {
                    for (int i = 0; i < data.columns(); i++) {
                        cstyles.put(HEADER + i, getCellStyle(getCell(sheet, startRow, startCol + i)));
                    }
                }
                int styleRow = startRow + (data.hasColumnHeader() ? 1 : 0);
                for (int i = 0; i < data.columns(); i++) {
                    Cell cell = getCell(sheet, styleRow, startCol + i);
                    cstyles.put(COLUMN + i, getCellStyle(cell));
                }
                break;
            case STYLE_NAME_PREFIX:
                if (data.hasColumnHeader()) {
                    for (int i = 0; i < data.columns(); i++) {
                        String prefix = styleNamePrefix + SEP + HEADER;
                        // Check for style <STYLE_NAME_PREFIX><SEP><HEADER><SEP><COLUMN_NAME>
                        CellStyle cs = getCellStyle(prefix + SEP + data.getColumnName(i));
                        // Check for style <STYLE_NAME_PREFIX><SEP><HEADER><SEP><COLUMN_INDEX>
                        if (cs == null)
                            cs = getCellStyle(prefix + SEP + (i + 1));
                        // Check for style <STYLE_NAME_PREFIX><SEP><HEADER>
                        if (cs == null)
                            cs = getCellStyle(prefix);
                        if (cs == null)
                            cs = new SSCellStyle(workbook, workbook.getCellStyleAt((short) 0));

                        cstyles.put(HEADER + i, cs);
                    }
                }
                for (int i = 0; i < data.columns(); i++) {
                    String prefix = styleNamePrefix + SEP + COLUMN;
                    // Check for style <STYLE_NAME_PREFIX><SEP><COLUMN><SEP><COLUMN_NAME>
                    CellStyle cs = getCellStyle(prefix + SEP + data.getColumnName(i));
                    // Check for style <STYLE_NAME_PREFIX><SEP><COLUMN><SEP><COLUMN_INDEX>
                    if (cs == null)
                        cs = getCellStyle(prefix + SEP + (i + 1));
                    // Check for style <STYLE_NAME_PREFIX><SEP><COLUMN><SEP><DATA_TYPE>
                    if (cs == null)
                        cs = getCellStyle(prefix + SEP + data.getColumnType(i).toString());
                    if (cs == null)
                        cs = new SSCellStyle(workbook, workbook.getCellStyleAt((short) 0));

                    cstyles.put(COLUMN + i, cs);
                }
                break;
            case DATA_FORMAT_ONLY:
                if (data.hasColumnHeader()) {
                    for (int i = 0; i < data.columns(); i++) {
                        cstyles.put(HEADER + i, DataFormatOnlyCellStyle.get(DataType.String));
                    }
                }
                for (int i = 0; i < data.columns(); i++) {
                    cstyles.put(COLUMN + i, DataFormatOnlyCellStyle.get(data.getColumnType(i)));
                }
                break;
            default:
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
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress cra = sheet.getMergedRegion(i);
            if (cra.formatAsString().equals(reference)) {
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
     * <p>
     * Reads the workbook if the file exists, otherwise creates a new workbook of the corresponding format.
     *
     * @param excelFile Microsoft Excel file to read or create if not existing
     * @return Instance of the workbook
     */
    public static Workbook getWorkbook(File excelFile, String password, boolean create) throws IOException {
        Workbook wb;

        if (excelFile.exists()) {
            if (password == null)
                wb = new Workbook(excelFile);
            else
                wb = new Workbook(excelFile, password);
        } else {
            if (create) {
                String filename = excelFile.getName().toLowerCase();
                if (filename.endsWith(".xls")) {
                    wb = new Workbook(excelFile, SpreadsheetVersion.EXCEL97);
                } else if (filename.endsWith(".xlsx")) {
                    wb = new Workbook(excelFile, SpreadsheetVersion.EXCEL2007);
                } else
                    throw new IllegalArgumentException("File extension \"" + filename.substring(filename.lastIndexOf('.') + 1) + "\" not supported! Only *.xls and *.xlsx are allowed!");
            } else
                throw new FileNotFoundException("File '" + excelFile.getName() + "' could not be found - " +
                        "you may specify to automatically create the file if not existing.");
        }
        return wb;
    }

    public static Workbook getWorkbook(File excelFile, boolean create) throws IOException {
        return getWorkbook(excelFile, null, create);
    }

    public static Workbook getWorkbook(String filename, String password, boolean create) throws IOException {
        return Workbook.getWorkbook(new File(filename), password, create);
    }

    public static Workbook getWorkbook(String filename, boolean create) throws IOException {
        return Workbook.getWorkbook(new File(filename), create);
    }

    public void setCellFormula(Cell c, String formula) {
        c.setCellFormula(formula);
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

    public int getLastRow(int sheetIndex) {
        return getSheet(sheetIndex).getLastRowNum();
    }

    public int getLastRow(String sheetName) {
        return getSheet(sheetName).getLastRowNum();
    }

    public int getLastColumn(Sheet sheet) {
        int lastRow = sheet.getLastRowNum();
        int lastColumn = 1;
        for (int i = 0; i < lastRow; ++i) {
            Row row = sheet.getRow(i);
            if (row != null) {
                int col = row.getLastCellNum();
                if (col > lastColumn) {
                    lastColumn = col;
                }
            }
        }
        return lastColumn - 1;
    }

    public int getLastColumn(int sheetIndex) {
        return getLastColumn(getSheet(sheetIndex));
    }

    public int getLastColumn(String sheetName) {
        return getLastColumn(getSheet(sheetName));
    }

    public void appendNamedRegion(DataFrame data, String name, String worksheetScope, boolean header, boolean overwriteFormulaCells) {
        Sheet sheet = workbook.getSheet(getName(name, worksheetScope).getSheetName());
        // top, left, bottom, right
        int[] coord = getReferenceCoordinatesForName(name, worksheetScope);
        writeData(data, sheet, coord[2] + 1, coord[1], header, overwriteFormulaCells);
        int bottom = coord[2] + data.rows();
        int right = Math.max(coord[1] + data.columns() - 1, coord[3]);
        CellRangeAddress cra = new CellRangeAddress(coord[0], bottom, coord[1], right);
        String formula = cra.formatAsString(sheet.getSheetName(), true);
        createName(name, worksheetScope, formula, true);
    }


    public void appendWorksheet(DataFrame data, int worksheetIndex, boolean header) {
        Sheet sheet = getSheet(worksheetIndex);
        int lastRow = getLastRow(worksheetIndex);
        int firstCol = Integer.MAX_VALUE;
        for (int i = 0; i < lastRow && firstCol > 0; i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getFirstCellNum() < firstCol)
                firstCol = row.getFirstCellNum();
        }
        if (firstCol == Integer.MAX_VALUE)
            firstCol = 0;

        writeWorksheet(data, worksheetIndex, getLastRow(worksheetIndex) + 1, firstCol, header, false);
    }

    public void appendWorksheet(DataFrame data, String worksheetName, boolean header) {
        appendWorksheet(data, workbook.getSheetIndex(worksheetName), header);
    }

    public void clearSheet(int sheetIndex) {
        Sheet sheet = getSheet(sheetIndex);
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        for (int i = lastRow; i >= firstRow; i--) {
            Row r = sheet.getRow(i);
            if (r != null)
                sheet.removeRow(r);
        }
    }

    public void clearSheet(String sheetName) {
        clearSheet(workbook.getSheetIndex(sheetName));
    }

    // coords[] = { top, left, bottom, right }
    public void clearRange(int sheetIndex, int[] coords) {
        Sheet sheet = getSheet(sheetIndex);
        for (int i = coords[0]; i <= coords[2]; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            for (int j = coords[1]; j <= coords[3]; j++) {
                Cell cell = row.getCell(j);
                if (cell != null)
                    row.removeCell(cell);
            }
            if (row.getLastCellNum() < 0)
                sheet.removeRow(row);
        }
    }

    public void clearRange(String sheetName, int[] coords) {
        clearRange(workbook.getSheetIndex(sheetName), coords);
    }

    public void clearRangeFromReference(String reference) {
        AreaReference ref = new AreaReference(reference, workbook.getSpreadsheetVersion());
        CellReference firstCell = ref.getFirstCell();
        CellReference lastCell = ref.getLastCell();
        String sheetName = firstCell.getSheetName();
        int[] coords = {firstCell.getRow(), firstCell.getCol(), lastCell.getRow(),
                lastCell.getCol()};
        clearRange(sheetName, coords);
    }

    public void clearNamedRegion(String name, String worksheetScope) {
        String dataSourceSheetName = getName(name, worksheetScope).getSheetName();
        int[] coords = getReferenceCoordinatesForName(name, worksheetScope);
        clearRange(dataSourceSheetName, coords);
    }

    public void createFreezePane(int sheetIndex, int colSplit, int rowSplit, int leftColumn, int topRow) {
        if (leftColumn < 0 | topRow < 0)
            getSheet(sheetIndex).createFreezePane(colSplit, rowSplit);
        else
            getSheet(sheetIndex).createFreezePane(colSplit, rowSplit, leftColumn, topRow);
    }

    public void createFreezePane(String sheetName, int colSplit, int rowSplit, int leftColumn, int topRow) {
        createFreezePane(workbook.getSheetIndex(sheetName), colSplit, rowSplit, leftColumn, topRow);
    }

    public void createFreezePane(int sheetIndex, int colSplit, int rowSplit) {
        createFreezePane(sheetIndex, colSplit, rowSplit, -1, -1);
    }

    public void createFreezePane(String sheetName, int colSplit, int rowSplit) {
        createFreezePane(sheetName, colSplit, rowSplit, -1, -1);
    }

    public void createSplitPane(int sheetIndex, int xSplitPos, int ySplitPos, int leftColumn, int topRow) {
        getSheet(sheetIndex).createSplitPane(xSplitPos, ySplitPos, leftColumn, topRow, PaneType.LOWER_RIGHT);
    }

    public void createSplitPane(String sheetName, int xSplitPos, int ySplitPos, int leftColumn, int topRow) {
        createSplitPane(workbook.getSheetIndex(sheetName), xSplitPos, ySplitPos, leftColumn, topRow);
    }

    public void removePane(int sheetIndex) {
        createFreezePane(sheetIndex, 0, 0);
    }

    public void removePane(String sheetName) {
        createFreezePane(sheetName, 0, 0);
    }

    public void setSheetColor(int sheetIndex, int color) {
        if (isXSSF()) {
            XSSFWorkbook wb = (XSSFWorkbook) workbook;
            XSSFSheet sheet = wb.getSheetAt(sheetIndex);
            sheet.setTabColor(
                    new XSSFColor(IndexedColors.fromInt(color), wb.getStylesSource().getIndexedColors()));
        } else if (isHSSF()) {
            this.warnings.add("Setting the sheet color for XLS files is not supported yet.");
        }
    }

    public void setSheetColor(String sheetName, int color) {
        setSheetColor(workbook.getSheetIndex(sheetName), color);
    }

    public int[] getBoundingBox(int sheetIndex, int startRow, int startCol, int endRow, int endCol,
                                boolean autofitRow, boolean autofitCol) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        final int mark = Integer.MAX_VALUE - 1;

        if (startRow < 0) {
            startRow = sheet.getFirstRowNum();
            if (sheet.getRow(startRow) == null) {
                // There is no row in this sheet
                startRow = -1;
            }
        }

        if (endRow < 0) {
            // We interpret this as "all except for the last N rows"
            // -1 => auto-detect the last row
            // -2 => all except for the last row
            // -3 => all except for the last 2 rows
            // ...
            endRow = sheet.getLastRowNum() + endRow + 1;
            if (sheet.getRow(endRow) == null) {
                // There is no row in this sheet
                endRow = -1;
            }
        }

        int minRow = startRow;
        int maxRow = endRow;
        int minCol = startCol;
        int maxCol = endCol < 0 ? mark : endCol;

        int origEndCol = endCol;

        startCol = startCol < 0 ? mark : startCol;
        endCol = endCol < 0 ? -1 : endCol;
        Cell topLeft = null, bottomRight = null;
        boolean anyCell = false;
        for (int i = minRow; i > -1 && i <= maxRow; i++) {
            Row r = sheet.getRow(i);
            if (r != null) {
                // Determine column boundaries
                int start = Math.max(minCol, r.getFirstCellNum());
                int end = Math.min(maxCol + 1, r.getLastCellNum()); // NOTE: getLastCellNum is 1-based!
                boolean anyNonBlank = false;
                for (int j = start; j > -1 && j < end; j++) {
                    Cell c = r.getCell(j);
                    if (c != null && c.getCellType() != CellType.BLANK) {
                        anyCell = true;
                        anyNonBlank = true;
                        if ((autofitCol || minCol < 0) && (topLeft == null || j < startCol)) {
                            startCol = j;
                            topLeft = c;
                        }
                        if ((autofitCol || maxCol == mark) && (bottomRight == null || j > endCol)) {
                            endCol = j;
                            bottomRight = c;
                        }
                    }
                }
                if (autofitRow && anyNonBlank) {
                    endRow = i;
                    if (sheet.getRow(startRow) == null) {
                        startRow = i;
                    }
                }
            }
        }

        if ((autofitRow || startRow < 0) && !anyCell) {
            startRow = endRow = -1;
        }
        if ((autofitCol || startCol == mark) && !anyCell) {
            startCol = endCol = -1;
        }

        if (origEndCol < 0) {
            // We interpret this as "all except for the last N columns"
            // -1 => auto-detect the last column
            // -2 => all except for the last column
            // -3 => all except for the last 2 columns
            // ...
            endCol = endCol + origEndCol + 1;
        }

        return new int[]{startRow, startCol, endRow, endCol};
    }

    public int[] getBoundingBox(String sheetName, int startRow, int startCol, int endRow, int endCol,
                                boolean autofitRow, boolean autofitColumn) {
        return getBoundingBox(workbook.getSheetIndex(sheetName), startRow, startCol, endRow, endCol,
                autofitRow, autofitColumn);
    }

    public List<String> getAndClearWarnings() {
        List<String> warnings = this.warnings;
        this.warnings = new ArrayList<>();
        return warnings;
    }
}
