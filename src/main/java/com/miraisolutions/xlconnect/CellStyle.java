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

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;

public interface CellStyle extends SupportsWarnings {
    void setBorderBottom(BorderStyle border);

    void setBorderLeft(BorderStyle border);

    void setBorderRight(BorderStyle border);

    void setBorderTop(BorderStyle border);

    void setBottomBorderColor(short color);

    void setLeftBorderColor(short color);

    void setRightBorderColor(short color);

    void setTopBorderColor(short color);

    void setDataFormat(String format);

    void setFillBackgroundColor(short bg);

    void setFillForegroundColor(short fp);

    void setFillPattern(FillPatternType bg);

    void setWrapText(boolean wrap);
}
