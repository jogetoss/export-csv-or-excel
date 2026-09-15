package org.joget.marketplace.util;

import java.io.*;
import java.nio.file.*;
import java.util.Properties;

/** Disk-backed job state, shared by the export worker and polling requests. */
public final class BackgroundExportStatus {
    private final File folder;
    private final boolean storeToForm;
    private long processed;
    private int total;

    public BackgroundExportStatus(File folder, boolean storeToForm) {
        this.folder = folder;
        this.storeToForm = storeToForm;
    }

    public void stage(String stage) throws IOException {
        update(stage, processed, total);
    }

    public void update(String stage, long processed, int total) throws IOException {
        this.processed = processed;
        this.total = total;
        Properties state = new Properties();
        state.setProperty("stage", stage);
        state.setProperty("processed", Long.toString(processed));
        state.setProperty("total", Integer.toString(total));
        state.setProperty("storeToForm", Boolean.toString(storeToForm));
        Path temporary = Files.createTempFile(folder.toPath(), "progress-", ".tmp");
        try {
            try (OutputStream out = Files.newOutputStream(temporary)) {
                state.store(out, null);
            }
            Path target = new File(folder, "progress.properties").toPath();
            try {
                Files.move(temporary, target, StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING);
            } catch (AtomicMoveNotSupportedException e) {
                Files.move(temporary, target, StandardCopyOption.REPLACE_EXISTING);
            }
        } finally {
            Files.deleteIfExists(temporary);
        }
    }

    public static String readJson(File folder, File exportFile) throws IOException {
        Properties state = new Properties();
        File statusFile = new File(folder, "progress.properties");
        if (statusFile.isFile()) {
            try (InputStream in = new FileInputStream(statusFile)) { state.load(in); }
        }
        String stage = state.getProperty("stage", "preparing");
        // Support exports started before the progress feature was installed.
        if (new File(exportFile.getPath() + ".completed").isFile()) { stage = "ready"; }
        if (!stage.matches("preparing|exporting|finalizing|storing|ready|failed")) { stage = "failed"; }
        long processed = Long.parseLong(state.getProperty("processed", "0"));
        long total = Long.parseLong(state.getProperty("total", "0"));
        long percent = total > 0 ? Math.min(99, Math.max(0, processed * 100 / total)) : 0;
        if ("ready".equals(stage)) { percent = 100; }
        return "{\"stage\":\"" + stage + "\",\"processed\":" + processed
                + ",\"total\":" + total + ",\"percent\":" + percent
                + ",\"storeToForm\":" + Boolean.parseBoolean(state.getProperty("storeToForm")) + "}";
    }
}
