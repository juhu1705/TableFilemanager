/*
 * ExcelAndCSVToArray
 * Config.java
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

import de.noisruker.config.ConfigElement;
import de.noisruker.config.ConfigElementType;
import de.noisruker.config.ConfigManager;

import java.io.IOException;

/**
 * Die Konfigurationen für den filemanager
 *
 * @author Fabius
 * @category Einstellungen (Import / Export)
 */
public class Config {

    /**
     * Ob ein output header oben in die Tabelle eingefügt werden soll
     */
    @ConfigElement(defaultValue = "false", type = ConfigElementType.CHECK, description = "hasHeaderOutput.description", name = "hasHeaderOutput.text", location = "config.export", visible = true)
    public static boolean hasHeaderOutput = true;

    /**
     * Register the Config values to the config manager. Use this if you want the header output to be configurable. Elsewhere, the header output is set to true.
     * @throws IOException If something went wrong due to the value registration. Post the error as issue in this repository
     */
    public static void init() throws IOException {
        ConfigManager.getInstance().register(Config.class);
    }

}
