/*
 * ExcelAndCSVToArray
 * CellStyles.java
 * Copyright © 2021 Fabius Mettner
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

package de.noisruker.filemanager.utils;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Beinhaltet alle Zell-Formatierungsformen, die benutzt werden.
 * 
 * @author Juhu1705
 * @category Import / Export
 */
public class CellStyles {

	public static HSSFCellStyle title(HSSFWorkbook workbook) {
		HSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setColor(HSSFColorPredefined.WHITE.getIndex());
		HSSFCellStyle style = workbook.createCellStyle();

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);

		font.setFontHeightInPoints((short) 16);

		style.setFont(font);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// (HSSFColorPredefined.RED.getIndex());
		style.setFillForegroundColor(HSSFColorPredefined.GREY_80_PERCENT.getIndex());
		return style;
	}

	public static HSSFCellStyle header(HSSFWorkbook workbook) {
		HSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		font.setFontHeightInPoints((short) 24);
		HSSFCellStyle style = workbook.createCellStyle();

		style.setFont(font);
		style.setAlignment(HorizontalAlignment.CENTER);

		return style;
	}

	public static HSSFCellStyle normal2(HSSFWorkbook workbook) {
		HSSFFont font = workbook.createFont();
		font.setBold(false);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		HSSFCellStyle style = workbook.createCellStyle();

		font.setFontHeightInPoints((short) 12);

		style.setFont(font);
		// style.setBorderBottom(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		// style.setBorderTop(BorderStyle.THIN);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// (HSSFColorPredefined.RED.getIndex());
		style.setFillForegroundColor(HSSFColorPredefined.WHITE.getIndex());
		return style;
	}

	public static HSSFCellStyle normal1(HSSFWorkbook workbook) {
		HSSFFont font = workbook.createFont();
		font.setBold(false);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		HSSFCellStyle style = workbook.createCellStyle();

		font.setFontHeightInPoints((short) 12);

		style.setFont(font);
		// style.setBorderBottom(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		// style.setBorderTop(BorderStyle.THIN);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// (HSSFColorPredefined.RED.getIndex());
		style.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
		return style;
	}

	public static HSSFCellStyle up(HSSFWorkbook workbook) {
		HSSFFont font = workbook.createFont();
		font.setBold(false);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		HSSFCellStyle style = workbook.createCellStyle();

		font.setFontHeightInPoints((short) 12);

		style.setFont(font);
		// style.setBorderBottom(BorderStyle.THIN);
		// style.setBorderRight(BorderStyle.THIN);
		// style.setBorderLeft(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// (HSSFColorPredefined.RED.getIndex());
		style.setFillForegroundColor(HSSFColorPredefined.WHITE.getIndex());
		return style;
	}

	public static XSSFCellStyle title(XSSFWorkbook workbook) {
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setColor(HSSFColorPredefined.WHITE.getIndex());
		XSSFCellStyle style = workbook.createCellStyle();

		font.setFontHeightInPoints((short) 16);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);

		style.setFont(font);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// (HSSFColorPredefined.RED.getIndex());
		style.setFillForegroundColor(HSSFColorPredefined.GREY_80_PERCENT.getIndex());
		return style;
	}

	public static XSSFCellStyle header(XSSFWorkbook workbook) {
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		font.setFontHeightInPoints((short) 24);

		XSSFCellStyle style = workbook.createCellStyle();

		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(font);

		return style;
	}

	public static XSSFCellStyle normal2(XSSFWorkbook workbook) {
		XSSFFont font = workbook.createFont();
		font.setBold(false);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		XSSFCellStyle style = workbook.createCellStyle();

		font.setFontHeightInPoints((short) 12);

		style.setFont(font);
		// style.setBorderBottom(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		// style.setBorderTop(BorderStyle.THIN);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// (HSSFColorPredefined.RED.getIndex());
		style.setFillForegroundColor(HSSFColorPredefined.WHITE.getIndex());
		return style;
	}

	public static XSSFCellStyle normal1(XSSFWorkbook workbook) {
		XSSFFont font = workbook.createFont();
		font.setBold(false);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		XSSFCellStyle style = workbook.createCellStyle();

		font.setFontHeightInPoints((short) 12);

		style.setFont(font);
		// style.setBorderBottom(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		// style.setBorderTop(BorderStyle.THIN);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// (HSSFColorPredefined.RED.getIndex());
		style.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
		return style;
	}

	public static XSSFCellStyle up(XSSFWorkbook workbook) {
		XSSFFont font = workbook.createFont();
		font.setBold(false);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		XSSFCellStyle style = workbook.createCellStyle();

		font.setFontHeightInPoints((short) 12);

		style.setFont(font);
		// style.setBorderBottom(BorderStyle.THIN);
		// style.setBorderRight(BorderStyle.THIN);
		// style.setBorderLeft(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// (HSSFColorPredefined.RED.getIndex());
		style.setFillForegroundColor(HSSFColorPredefined.WHITE.getIndex());
		return style;
	}
}
