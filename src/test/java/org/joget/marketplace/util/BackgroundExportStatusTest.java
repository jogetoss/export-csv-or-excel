package org.joget.marketplace.util;

import java.io.File;
import java.nio.file.Files;
import junit.framework.TestCase;

public class BackgroundExportStatusTest extends TestCase {
    public void testProgressAndCompletion() throws Exception {
        File folder = Files.createTempDirectory("export-status-test").toFile();
        File report = new File(folder, "report.xlsx");
        try {
            BackgroundExportStatus status = new BackgroundExportStatus(folder, true);
            status.update("exporting", 500, 1000);
            String json = BackgroundExportStatus.readJson(folder, report);
            assertTrue(json.contains("\"percent\":50"));
            assertTrue(json.contains("\"storeToForm\":true"));
            status.update("finalizing", 1000, 1000);
            assertTrue(BackgroundExportStatus.readJson(folder, report).contains("\"percent\":99"));
            status.stage("failed");
            assertTrue(BackgroundExportStatus.readJson(folder, report).contains("\"stage\":\"failed\""));
            status.stage("ready");
            assertTrue(BackgroundExportStatus.readJson(folder, report).contains("\"percent\":100"));
            status.update("exporting", 0, 0);
            assertTrue(BackgroundExportStatus.readJson(folder, report).contains("\"percent\":0"));
            new File(report.getPath() + ".completed").createNewFile();
            assertTrue(BackgroundExportStatus.readJson(folder, report).contains("\"stage\":\"ready\""));
        } finally {
            for (File file : folder.listFiles()) { Files.delete(file.toPath()); }
            Files.delete(folder.toPath());
        }
    }
}
