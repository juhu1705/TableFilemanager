/*
 * ExcelAndCSVToArray
 * WriteableContent.java
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

package de.noisruker.filemanager;

import de.noisruker.filemanager.utils.CellStyles;
import de.noisruker.filemanager.utils.Config;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;
import java.util.logging.Level;

import static de.noisruker.logger.Logger.LOGGER;
import static org.apache.poi.ss.usermodel.CellType.STRING;

/**
 * Diese Klasse dient als Schnittstelle zum Im- / Export. Ein Writeable Content stellt einen Sheet dar.
 * 
 * @author Juhu1705
 * @category Import / Export
 */
public class WriteableContent {

	private HashMap<Vec2i, String> lines = new HashMap<Vec2i, String>();
	private String name;

	private int startTable;

	public WriteableContent() {
		this.name = "";
	}

	public WriteableContent(String name) {
		this.name = name;
	}

	/**
	 * @param line In welcher Zeile gestartet werden soll
	 */
	public void setStartTable(int line) {
		this.startTable = line;
	}

	public String getName() {
		return this.name;
	}

	public WriteableContent setName(String name) {
		this.name = name;
		return this;
	}

	/**
	 * Schreibt daten in eine Tabellenzelle
	 * @param position Die Position der Tabellenzelle
	 * @param content Der Wert, auf den die Tabellenzelle gesetzt werden soll
	 * @return Dieses Element
	 */
	public WriteableContent addCell(Vec2i position, String content) {
		lines.put(position, content);
		return this;
	}

	/**
	 * Löscht den Inhalt der entsprechenden Zelle
	 * @param position Die Position der Tabellenzelle
	 * @return Dieses Element
	 */
	public WriteableContent removeCell(Vec2i position) {
		lines.remove(position);
		return this;
	}

	/**
	 * Fügt eine Zeile an der Startposition ein. Dabei ist die gesetzte Position die Zelle ganz Links in der Liste und alle weiteren Elemente werden rechts davon eingefügt.
	 * @param startPosition Die Startposition zum Einfügen.
	 * @param contents Die Inhalte die Einzufügen sind
	 * @return Dieses Element
	 */
	public WriteableContent addLine(Vec2i startPosition, String... contents) {
		for (int x = startPosition.x, i = 0; i < contents.length; x++, i++) {
			lines.put(new Vec2i(x, startPosition.y), contents[i]);
		}
		return this;
	}

	/**
	 * Fügt eine ArrayTabelle in diesen Writeable Content ein. Dabei ist die Startposition das Element links oben und alle weiteren Zeilen werden darunter eingefügt.
	 * @param startPosition Die Startposition zum Einfügen.
	 * @param grid Das 2dim Array das hinzugefügt werden soll
	 * @return Dieses Element
	 */
	public WriteableContent addGrid(Vec2i startPosition, String[]... grid) {
		for (int x = startPosition.x, i = 0; i < grid.length; x++, i++) {
			this.addLine(new Vec2i(x, startPosition.y), grid[i]);
		}
		return this;
	}

	/**
	 * Fügt eine Zeile an der Startposition ein. Dabei ist die gesetzte Position die Zelle ganz Links in der Liste und alle weiteren Elemente werden rechts davon eingefügt.
	 * @param startPosition Die Startposition zum Einfügen.
	 * @param contents Die Inhalte die Einzufügen sind
	 * @return Dieses Element
	 */
	public WriteableContent addLine(Vec2i startPosition, List<String> contents) {
		for (int x = startPosition.x, i = 0; i < contents.size(); x++, i++) {
			lines.put(new Vec2i(x, startPosition.y), contents.get(i));
		}
		return this;
	}

	/**
	 * Fügt eine ArrayTabelle in diesen Writeable Content ein. Dabei ist die Startposition das Element links oben und alle weiteren Zeilen werden darunter eingefügt.
	 * @param startPosition Die Startposition zum Einfügen.
	 * @param contents Das 2dim Array das hinzugefügt werden soll
	 * @return Dieses Element
	 */
	public WriteableContent addListGrid(Vec2i startPosition, List<List<String>> contents) {
		for (int x = startPosition.x, i = 0; i < contents.size(); x++, i++) {
			this.addLine(new Vec2i(x, startPosition.y), contents.get(i));
		}
		return this;
	}

	public WriteableContent removeLine(Vec2i startPosition, int length) {
		for (int x = startPosition.x, i = 0; i < length; x++, i++) {
			lines.remove(new Vec2i(x, startPosition.y));
		}
		return this;
	}

	public String[][] getReverseGrid() {
		Vec2i maxLength = this.getMaxLength();
		String[][] cells = new String[maxLength.x + 1][maxLength.y + 1];
		for (Entry<Vec2i, String> cell : lines.entrySet()) {
			cells[cell.getKey().x][cell.getKey().y] = cell.getValue();
		}
		return cells;
	}

	public String[][] getGrid() {
		Vec2i maxLength = this.getMaxLength();
		String[][] cells = new String[maxLength.y + 1][maxLength.x + 1];
		for (Entry<Vec2i, String> cell : lines.entrySet()) {
			cells[cell.getKey().y][cell.getKey().x] = cell.getValue();
		}
		return cells;
	}

	public String getStringAt(Vec2i position) {
		return this.lines.get(position);
	}

	public void writeXLS(HSSFWorkbook workbook, HSSFSheet sheet, int startingLineY) {
		Vec2i maxLength = this.getMaxLength();
		String[][] cells = this.getReverseGrid();

		int rownumber = startingLineY;

		Row row;

		Cell cell;

		for (int y = 0; y < this.startTable; y++) {
			HSSFCellStyle style = CellStyles.header(workbook);
			row = sheet.createRow(y);
			for (int x = 0; x <= maxLength.x; x++) {
				cell = row.createCell(x, STRING);
				cell.setCellValue(cells[x][0] == null ? "" : cells[x][y]);
				cell.setCellStyle(style);
			}

		}

		HSSFCellStyle style = CellStyles.title(workbook);
		row = sheet.createRow(this.startTable);
		for (int x = 0; x <= maxLength.x; x++) {
			cell = row.createCell(x, STRING);
			cell.setCellValue(cells[x][this.startTable] == null ? "" : cells[x][this.startTable]);
			cell.setCellStyle(style);

		}

		style = CellStyles.normal2(workbook);
		row = sheet.createRow(this.startTable + 1);
		for (int x = 0; x <= maxLength.x; x++) {
			cell = row.createCell(x, STRING);
			cell.setCellValue("");
			cell.setCellStyle(style);
		}

		for (int y = this.startTable + 1; y <= maxLength.y; y++) {
			if (y % 2 == 1) {
				style = CellStyles.normal2(workbook);
				row = sheet.createRow(rownumber + y);
				for (int x = 0; x <= maxLength.x; x++) {
					cell = row.createCell(x, STRING);
					cell.setCellValue(cells[x][y] == null ? "" : cells[x][y]);
					cell.setCellStyle(style);
				}
			} else {
				style = CellStyles.normal1(workbook);
				row = sheet.createRow(rownumber + y);
				for (int x = 0; x <= maxLength.x; x++) {
					cell = row.createCell(x, STRING);
					cell.setCellValue(cells[x][y] == null ? "" : cells[x][y]);
					cell.setCellStyle(style);
				}
			}
		}

		style = CellStyles.up(workbook);
		row = sheet.createRow(maxLength.y + 1 + rownumber);
		for (int x = 0; x <= maxLength.x; x++) {
			cell = row.createCell(x, STRING);
			cell.setCellValue("");
			cell.setCellStyle(style);
		}

		if (Config.hasHeaderOutput)
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, maxLength.x));

		for (int y = 0; y <= maxLength.y; y++)
			sheet.autoSizeColumn(y);

	}

	public void writeCSV(BufferedWriter writer) {
		Vec2i maxLength = this.getMaxLength();
		String[][] cells = this.getReverseGrid();

		for (int y = 0; y <= maxLength.y; y++) {
			try {
				for (int x = 0; x <= maxLength.x; x++) {
					if (cells[x][y] != null)
						writer.append(cells[x][y] + ";");

				}
				writer.newLine();
			} catch (IOException e) {
				LOGGER.log(Level.SEVERE, "Failed to write CSV", e);
			}
		}
	}

	public void writeXLSX(XSSFWorkbook workbook, XSSFSheet sheet, int startingLineY) {
		Vec2i maxLength = this.getMaxLength();
		String[][] cells = this.getReverseGrid();

		int rownumber = startingLineY;

		Row row;

		Cell cell;

		for (int y = 0; y < this.startTable; y++) {
			XSSFCellStyle style = CellStyles.header(workbook);
			row = sheet.createRow(y);
			for (int x = 0; x <= maxLength.x; x++) {
				cell = row.createCell(x, STRING);
				cell.setCellValue(cells[x][0] == null ? "" : cells[x][y]);
				cell.setCellStyle(style);
			}

		}

		XSSFCellStyle style = CellStyles.title(workbook);
		row = sheet.createRow(this.startTable);
		for (int x = 0; x <= maxLength.x; x++) {
			cell = row.createCell(x, STRING);
			cell.setCellValue(cells[x][this.startTable] == null ? "" : cells[x][this.startTable]);
			cell.setCellStyle(style);

		}

		style = CellStyles.normal2(workbook);
		row = sheet.createRow(this.startTable + 1);
		for (int x = 0; x <= maxLength.x; x++) {
			cell = row.createCell(x, STRING);
			cell.setCellValue("");
			cell.setCellStyle(style);
		}

		for (int y = this.startTable + 1; y <= maxLength.y; y++) {
			if (y % 2 == 1) {
				style = CellStyles.normal2(workbook);
				row = sheet.createRow(rownumber + y);
				for (int x = 0; x <= maxLength.x; x++) {
					cell = row.createCell(x, STRING);
					cell.setCellValue(cells[x][y] == null ? "" : cells[x][y]);
					cell.setCellStyle(style);
				}
			} else {
				style = CellStyles.normal1(workbook);
				row = sheet.createRow(rownumber + y);
				for (int x = 0; x <= maxLength.x; x++) {
					cell = row.createCell(x, STRING);
					cell.setCellValue(cells[x][y] == null ? "" : cells[x][y]);
					cell.setCellStyle(style);
				}
			}
		}

		style = CellStyles.up(workbook);
		row = sheet.createRow(maxLength.y + 1 + rownumber);
		for (int x = 0; x <= maxLength.x; x++) {
			cell = row.createCell(x, STRING);
			cell.setCellValue("");
			cell.setCellStyle(style);
		}

		if (Config.hasHeaderOutput)
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, maxLength.x));

		for (int y = 0; y <= maxLength.y; y++)
			sheet.autoSizeColumn(y);

	}

	private Vec2i getMaxLength() {
		Vec2i maxLength = new Vec2i();
		for (Entry<Vec2i, String> cell : lines.entrySet()) {
			maxLength.x = cell.getKey().x > maxLength.x ? cell.getKey().x : maxLength.x;
			maxLength.y = cell.getKey().y > maxLength.y ? cell.getKey().y : maxLength.y;
		}
		return maxLength;
	}

}
