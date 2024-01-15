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

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

public final class DataFrame {

    private final ArrayList<String> columnNames;
    private final ArrayList<Column> columns;

    public DataFrame() {
        this.columnNames = new ArrayList<>();
        this.columns = new ArrayList<>();
    }

    public int columns() {
        return columns.size();
    }

    public int rows() {
        return isEmpty() ? 0 : columns.get(0).size();
    }

    public boolean isEmpty() {
        return columns.isEmpty();
    }

    public boolean hasColumnHeader() {
        return columnNames.stream().anyMatch(Objects::nonNull);
    }


    public void addColumn(String name, Column column) {
        if (isEmpty() || (column.size() == rows())) {
            columnNames.add(name);
            columns.add(column);
        } else
            throw new IllegalArgumentException("Length of specified column does not match length of " +
                    "existing columns in the DataFrame!");
    }

    public String getColumnName(int index) {
        return columnNames.get(index);
    }

    public DataType getColumnType(int index) {
        return columns.get(index).getDataType();
    }

    public Column getColumn(int index) {
        return columns.get(index);
    }

    public List<String> getColumnNames() {
        return columnNames;
    }

    public List<DataType> getColumnTypes() {
        return columns.stream()
                .map(Column::getDataType)
                .collect(Collectors.toList());
    }
}
