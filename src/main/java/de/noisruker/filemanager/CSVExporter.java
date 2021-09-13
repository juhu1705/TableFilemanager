/*
 * ExcelAndCSVToArray
 * CSVExporter.java
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

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.logging.Level;

import static de.noisruker.logger.Logger.LOGGER;

/**
 * Verwaltet das Exportieren eines {@link WriteableContent} nach CSV an den
 * gegebenen Pfad.
 * 
 * @author Juhu1705
 * @category Export
 */
public class CSVExporter {

	protected CSVExporter() {
	}

	public static void writeCSV(String pathfile, WriteableContent toWrite) throws IOException {

		FileWriter fileWriter = null;
		BufferedWriter writer = null;

		try {
			fileWriter = new FileWriter(new File(pathfile + ".csv"));
			writer = new BufferedWriter(fileWriter);
		} catch (IOException e) {
			LOGGER.log(Level.SEVERE, "Fehler beim Erstellen einer .csv Datei", e);
		}

		toWrite.writeCSV(writer);

		writer.close();

	}

	public static void writeCSV(String pathfile, WriteableContent... toWrite) throws IOException {

		FileWriter fileWriter = null;
		BufferedWriter writer = null;

		try {
			fileWriter = new FileWriter(new File(pathfile + ".csv"));
			writer = new BufferedWriter(fileWriter);
		} catch (IOException e) {
			LOGGER.log(Level.SEVERE, "Fehler beim Erstellen einer .csv Datei", e);
		}

		for (WriteableContent writeable : toWrite) {
			writeable.writeCSV(writer);
			try {
				writer.newLine();
				writer.newLine();
			} catch (IOException e) {
				LOGGER.log(Level.WARNING, "Exception caused while exporting data: ", e);
			}

		}

		writer.close();

	}

}
