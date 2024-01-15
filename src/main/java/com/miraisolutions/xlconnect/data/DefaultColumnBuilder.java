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

package com.miraisolutions.xlconnect.data;

import com.miraisolutions.xlconnect.ErrorBehavior;
import com.miraisolutions.xlconnect.utils.CellUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


public final class DefaultColumnBuilder extends ColumnBuilder {

    // The following split is done for performance reasons
    private final String[] missingValueStrings;
    private final double[] missingValueNumbers;

    public DefaultColumnBuilder(int nrows, boolean forceConversion,
                                boolean takeCached, FormulaEvaluator evaluator, ErrorBehavior onErrorCell,
                                Object[] missingValue, String dateTimeFormat) {

        super(nrows, forceConversion, takeCached, evaluator, onErrorCell, dateTimeFormat);

        // Split missing values into missing values for strings and doubles
        // (for better performance later on)
        Map<Boolean, List<Object>> partitioned = Arrays.stream(missingValue)
                .collect(Collectors.partitioningBy(o -> o instanceof String));

        missingValueStrings = partitioned.get(true).toArray(String[]::new);
        missingValueNumbers = partitioned.get(false).stream().mapToDouble(o -> (Double) o).toArray();
    }

    @Override
    protected void handleCell(Cell c, CellValue cv) {
        // Determine (evaluated) cell data type
        switch (cv.getCellType()) {
            case BLANK:
                addMissing();
                break;
            case BOOLEAN:
                addValue(c, cv, DataType.Boolean);
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c))
                    addValue(c, cv, DataType.DateTime);
                else {
                    boolean missing = false;
                    for (double missingValueNumber : missingValueNumbers) {
                        if (cv.getNumberValue() == missingValueNumber) {
                            missing = true;
                            break;
                        }
                    }
                    if (missing)
                        addMissing();
                    else
                        addValue(c, cv, DataType.Numeric);
                }
                break;
            case STRING:
                boolean missing = false;
                for (String missingValueString : missingValueStrings) {
                    String value = cv.getStringValue();
                    if (value == null || value.equals(missingValueString)) {
                        missing = true;
                        break;
                    }
                }
                if (missing)
                    addMissing();
                else
                    addValue(c, cv, DataType.String);
                break;
            case FORMULA:
                cellError("Formula detected in already evaluated cell " + CellUtils.formatAsString(c) + "!");
                break;
            case ERROR:
                cellError("Error detected in cell " + CellUtils.formatAsString(c) + " - " + CellUtils.getErrorMessage(cv.getErrorValue()));
                break;
            default:
                cellError("Unexpected cell type detected for cell " + CellUtils.formatAsString(c) + "!");
        }
    }
}
