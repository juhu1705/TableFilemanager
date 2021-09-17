module de.noisruker.tablefilemanager {
    requires de.noisruker.logger;
    requires de.noisruker.config;
    requires java.logging;
    requires org.apache.poi.ooxml;

    exports de.noisruker.filemanager;
    exports de.noisruker.filemanager.utils;
}