/*
 * ExcelAndCSVToArray
 * CSVImporter.java
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

import de.noisruker.logger.Logger;

import java.io.*;
import java.net.URISyntaxException;
import java.util.logging.Level;

/**
 * Importiert eine CSV Datei aus dem mitgegebenen Pfad.
 * 
 * 
 * @author Juhu1705
 * @category Import
 */
public class CSVImporter {

	protected CSVImporter() {
	}

	public static WriteableContent readCSV(String pathfile) throws IOException, URISyntaxException {
		WriteableContent writeable = new WriteableContent();

		InputStreamReader fileReader = new InputStreamReader(getInput(pathfile), "UTF8");
		BufferedReader reader = new BufferedReader(fileReader);

		String zeile = "";
		int y = 0, x;

		while (reader.ready() && (zeile = reader.readLine()) != null) {
			x = 0;

			String parameter = "";
			for (char c : zeile.toCharArray()) {
				if (c == ';') {
					writeable.addCell(new Vec2i(x++, y), parameter);
					parameter = "";
				} else
					parameter += c;
			}
			writeable.addCell(new Vec2i(x++, y), parameter);
			y++;
		}

		reader.close();

		return writeable;
	}

	private static InputStream getInput(String name) throws URISyntaxException, FileNotFoundException {
		InputStream output;
		output = ExcelImporter.class.getClassLoader().getResourceAsStream(name);

		if (output == null) {
			try {
				output = new FileInputStream(new File(name));
			} catch (FileNotFoundException e) {
				Logger.LOGGER.log(Level.SEVERE, "", e);
			}
		}

		return output;
	}

}
