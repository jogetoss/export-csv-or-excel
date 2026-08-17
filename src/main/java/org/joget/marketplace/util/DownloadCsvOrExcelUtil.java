package org.joget.marketplace.util;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListColumn;
import org.joget.apps.datalist.model.DataListColumnFormat;
import org.joget.apps.datalist.model.DataListCollection;
import org.joget.apps.datalist.service.DataListService;
import org.joget.apps.form.model.FormRow;
import org.joget.commons.util.FileManager;
import org.joget.commons.util.LogUtil;
import org.joget.commons.util.SecurityUtil;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import java.io.ByteArrayOutputStream;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.URLEncoder;
import java.util.Collection;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.util.regex.Pattern;

import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.tika.Tika;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppService;
import org.joget.apps.datalist.lib.BeanShellColumn;
import org.joget.apps.datalist.model.DataListDisplayColumnProxy;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.service.FileUtil;
import org.joget.apps.form.service.FormUtil;
import org.joget.commons.util.UuidGenerator;
import org.springframework.context.ApplicationContext;

public class DownloadCsvOrExcelUtil {

    /*
     * Streaming export settings.
     *
     * DATA_BATCH_SIZE controls how many records are fetched from the datalist
     * binder at one time. SXSSF_ROW_WINDOW controls how many Excel rows Apache
     * POI retains in heap before flushing older rows to its temporary files.
     * These deliberately remain separate so they can be tuned independently
     * after the one-million-row performance test.
     */
    public static final int DATA_BATCH_SIZE = 2000;
    public static final int SXSSF_ROW_WINDOW = 100;
    private static final int XLSX_MAX_ROWS_PER_SHEET = 1048576;
    private static final int FILE_COPY_BUFFER_SIZE = 64 * 1024;
    private static final Pattern HTML_TAG_PATTERN = Pattern.compile("<[^>]*>");

    private final static DuplicateAndSkip duplicates = new DuplicateAndSkip();

    private static Map<String, Object> data;

    private final static String MESSAGE_PATH = "messages/DownloadCSVOrExcelDatalistAction";

    private static final Map<Workbook, CellStyle> NUMERIC_STYLE_CACHE = new java.util.WeakHashMap<>();
    private static final Map<Workbook, DataFormat> DATA_FORMAT_CACHE = new java.util.WeakHashMap<>();

    private static CellStyle getNumericStyle(Workbook wb) {
        CellStyle style = NUMERIC_STYLE_CACHE.get(wb);
        if (style != null) {
            return style;
        }

        DataFormat fmt = DATA_FORMAT_CACHE.computeIfAbsent(wb, k -> k.createDataFormat());

        style = wb.createCellStyle();
        style.setDataFormat(fmt.getFormat("#,##0.00"));
        NUMERIC_STYLE_CACHE.put(wb, style);
        return style;
    }

    /**
     * Streaming exports create one style per workbook. They do not use the
     * legacy static WeakHashMap caches, which are shared across export threads.
     */
    private static CellStyle createStreamingNumericStyle(Workbook workbook) {
        DataFormat format = workbook.createDataFormat();
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(format.getFormat("#,##0.00"));
        return style;
    }

    public static void storeCSVToForm(HttpServletRequest request, DataList dataList, DataListCollection dataListRows, String[] rowKeys, String renameFile, String fileName, String formDefId, String fileFieldId, String delimiter, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator,
            String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportEncrypt) {
        try {
            String csvFileName = renameFile.equalsIgnoreCase("true") ? fileName + ".csv" : "report.csv";

            File csvFile = createCsvFileForStorage(request, dataList, dataListRows, rowKeys, csvFileName, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportEncrypt);
            storeGeneratedFile(csvFile, formDefId, fileFieldId);
            csvFile.delete();
        } catch (IOException e) {
            LogUtil.error(getClassName(), e, "Failed to store CSV to form");
        }
    }

    public static void storeExcelToForm(Workbook workbook, String filename, String renameFile, String formDefId, String fileFieldId) {
        try {
            String excelFileName = renameFile.equalsIgnoreCase("true") ? filename : "report.xlsx";

            File excelFile = createExcelFileForStorage(workbook, excelFileName);
            storeGeneratedFile(excelFile, formDefId, fileFieldId);
            excelFile.delete();
        } catch (IOException e) {
            LogUtil.error(getClassName(), e, "Failed to store Excel to form");
        }
    }

    protected static File createCsvFileForStorage(HttpServletRequest request, DataList dataList,
            DataListCollection dataListRows, String[] rowKeys, String filename, String delimiter,
            String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator,
            String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportEncrypt) throws IOException {
        File csvFile = new File(FileManager.getBaseDirectory(), filename);
        try (PrintWriter writer = new PrintWriter(new OutputStreamWriter(new FileOutputStream(csvFile)))) {
            streamCSV(request, null, writer, dataList, dataListRows, rowKeys, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportEncrypt);
        }
        return csvFile;
    }

    protected static File createExcelFileForStorage(Workbook workbook, String filename) throws IOException {
        File excelFile = new File(FileManager.getBaseDirectory(), filename);
        try (FileOutputStream fileOut = new FileOutputStream(excelFile)) {
            workbook.write(fileOut);
        }
        return excelFile;
    }

    protected static void storeGeneratedFile(File generatedFile, String formDefId, String fileFieldId) {
        try {
            AppService appService = (AppService) FormUtil.getApplicationContext().getBean("appService");
            AppDefinition appDef = AppUtil.getCurrentAppDefinition();

            String recordId = UuidGenerator.getInstance().getUuid();
            String tableName = appService.getFormTableName(appDef, formDefId);

            FileUtil.storeFile(generatedFile, tableName, recordId);

            FormRowSet rows = new FormRowSet();
            FormRow row = new FormRow();
            row.setId(recordId);
            row.put(fileFieldId, generatedFile.getName());
            rows.add(row);

            appService.storeFormData(formDefId, tableName, rows, recordId);
        } catch (Exception e) {
            LogUtil.error(getClassName(), e, "Failed to store the generated file in the form.");
        }
    }

    public static File generateCSVFile(DataList dataList, DataListCollection dataListRows, String[] rowKeys, String renameFile, String fileName, String delimiter, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator, String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportEncrypt) throws Exception {
        StringWriter stringWriter = new StringWriter();
        PrintWriter writer = new PrintWriter(stringWriter);
        if (delimiter.isEmpty()) {
            delimiter = ",";
        }

        streamCSV(
                null, null,
                writer,
                dataList,
                dataListRows,
                rowKeys,
                delimiter,
                headerDecorator,
                downloadAllWhenNoneSelected,
                footerDecorator,
                includeCustomHeader,
                footerHeader,
                includeCustomFooter,
                exportEncrypt
        );

        writer.flush();
        String csvContent = stringWriter.toString();

        File outFile = generateCSVOutputFile(csvContent, fileName);

        return outFile;
    }

    public static void downloadCSV(HttpServletRequest request, HttpServletResponse response, DataList dataList,
            DataListCollection dataListRows, String[] rowKeys, String renameFile, String fileName, String delimiter,
            String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator,
            String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportEncrypt)
            throws ServletException, IOException {
        String filename = renameFile.equalsIgnoreCase("true") ? fileName + ".csv" : "report.csv";
        if (delimiter.isEmpty()) {
            delimiter = ",";
        }
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=" + filename + "");

        try (OutputStream outputStream = response.getOutputStream()) {
            PrintWriter writer = new PrintWriter(outputStream);
            streamCSV(request, response, writer, dataList, dataListRows, rowKeys, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportEncrypt);
            writer.flush(); // Flush any remaining buffered data
            outputStream.flush(); // Flush the output stream
            writer.close();
        }
    }

    protected static void streamCSV(HttpServletRequest request, HttpServletResponse response, PrintWriter writer, DataList dataList, DataListCollection dataListRows, String[] rowKeys, String delimiter, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator, String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportEncrypt) throws IOException {
        HashMap<String, StringBuilder> labelAndKeys = getLabelAndKey(dataList);
        StringBuilder keySB = labelAndKeys.get("key");
        StringBuilder headerSB = labelAndKeys.get("header");

        String[] keys = keySB.toString().split(",", 0);
        duplicates.setMap(findDuplicate(keys));

        if (includeCustomHeader(includeCustomHeader)) {
            writer.write((headerDecorator + "\n"));
        }

        if (delimiter != null && !delimiter.isEmpty()) {
            String replacedString = headerSB.toString().replace(",", delimiter);
            headerSB.setLength(0);
            headerSB.append(replacedString);
        }

        writer.write((headerSB + ""));

        if (rowKeys != null && rowKeys.length > 0) {
            if (!dataList.isUseSession()) {
                // rows already filtered by primary key IN (...); skip "id" matching
                for (int x = 0; x < dataListRows.size(); x++) {
                    Object row = getRow(dataListRows, x);
                    writeCSVContents(dataList, null, keys, row, writer, delimiter, exportEncrypt);
                }
            } else {
                //goes through all the datalist row
                for (int x = 0; x < dataListRows.size(); x++) {
                    //compare with all the rowkeys that have been selected
                    for (String rowKey : rowKeys) {

                        //check instance of HashMap if not it will be Formrow
                        boolean boolInstance = dataListRows.get(x) instanceof HashMap;
                        boolean foundRowKey = foundRowKey(boolInstance, dataListRows, x, rowKey);

                        //if no row is found skip
                        if (!foundRowKey) {
                            continue;
                        }

                        Object row = getRow(dataListRows, x);

                        //get the keys and save it
                        writeCSVContents(dataList, null, keys, row, writer, delimiter, exportEncrypt);
                        break;
                    }
                }
            }

        } else if (downloadAllWhenNoneSelected.equals("true")) {
            for (int x = 0; x < dataListRows.size(); x++) {
                Object row = getRow(dataListRows, x);
                //get the keys and save it
                writeCSVContents(dataList, null, keys, row, writer, delimiter, exportEncrypt);
            }
        }
        if (getFooter(footerHeader)) {
            writer.write("\n");
            writer.write((headerSB + "\n"));
        }
        if (includeCustomFooter(includeCustomFooter)) {
            writer.write("\n");
            writer.write((footerDecorator + "\n"));
        }
    }

    protected static void writeCSVContents(DataList dataList, ByteArrayOutputStream outputStream, String[] keys, Object row, PrintWriter writer, String delimiter, String exportEncrypt) throws IOException {
        // Construct CSV content
        StringBuilder stringBuilder = new StringBuilder();
        for (String value : keys) {
            String formattedValue = getBinderFormattedValue(dataList, row, value, null, exportEncrypt);

            if (formattedValue != null && formattedValue.contains(delimiter)) {
                formattedValue = "\"" + formattedValue + "\"";
            }

            stringBuilder.append(formattedValue);
            stringBuilder.append(delimiter);
        }

        // Remove the trailing delimiter if it exists
        if (stringBuilder.length() > 0 && stringBuilder.lastIndexOf(delimiter) == stringBuilder.length() - delimiter.length()) {
            stringBuilder.setLength(stringBuilder.length() - delimiter.length());
        }

        String value = stringBuilder.toString();

        // Write original CSV content to the output stream
        writer.write("\r\n");
        writer.write(value);
        writer.flush();

    }

    // ---------------------------------------------------------------------
    // New bounded-memory CSV export path
    // ---------------------------------------------------------------------

    /**
     * Writes CSV rows incrementally to disk. The legacy generateCSVFile(...)
     * method is preserved above, but it builds the full CSV in a StringWriter.
     */
    public static File generateStreamingCSVFile(DataList dataList, DataListCollection selectedRows, String[] rowKeys, File requestedFile, boolean makeUnique, String delimiter, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator, String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportEncrypt) throws IOException {
        long exportStartedAt = System.currentTimeMillis();
        int totalRows = getExpectedExportRows(dataList, selectedRows, rowKeys, downloadAllWhenNoneSelected);
        File outputFile = makeUnique ? getUniqueFile(requestedFile.getPath()) : requestedFile;
        File parent = outputFile.getParentFile();
        if (parent != null && !parent.isDirectory() && !parent.mkdirs()) {
            throw new IOException("Unable to create export directory: " + parent);
        }
        String actualDelimiter = delimiter == null || delimiter.isEmpty() ? "," : delimiter;
        LogUtil.info(getClassName(), getExportStartMessage("CSV", totalRows, outputFile));

        boolean completed = false;
        long processedRows = 0;
        try (PrintWriter writer = new PrintWriter(new OutputStreamWriter(new BufferedOutputStream(new FileOutputStream(outputFile), FILE_COPY_BUFFER_SIZE), "UTF-8"))) {
            List<DataListColumn> columns = getExportColumns(dataList);
            if (includeCustomHeader(includeCustomHeader)) {
                writer.write(headerDecorator);
                writer.write("\n");
            }
            writeStreamingCSVHeader(writer, columns, actualDelimiter);

            if (rowKeys != null && rowKeys.length > 0) {
                Set<String> selectedKeySet = new HashSet<>(Arrays.asList(rowKeys));
                if (selectedRows != null) {
                    for (int i = 0; i < selectedRows.size(); i++) {
                        if (selectedKeySet.contains(findRowKey(selectedRows, i))) {
                            writeStreamingCSVRow(writer, dataList, getRow(selectedRows, i), columns, actualDelimiter, exportEncrypt);
                            processedRows++;
                        }
                    }
                    logExportBatch("CSV", 1, selectedRows.size(), processedRows, totalRows, exportStartedAt);
                }
            } else if ("true".equals(downloadAllWhenNoneSelected)) {
                int start = 0;
                int batchNumber = 0;
                while (true) {
                    DataListCollection batch = dataList.getRows(DATA_BATCH_SIZE, start);
                    if (batch == null || batch.isEmpty()) {
                        break;
                    }
                    for (int i = 0; i < batch.size(); i++) {
                        writeStreamingCSVRow(writer, dataList, getRow(batch, i), columns, actualDelimiter, exportEncrypt);
                    }
                    int fetched = batch.size();
                    batchNumber++;
                    processedRows += fetched;
                    start += fetched;
                    logExportBatch("CSV", batchNumber, fetched, processedRows, totalRows, exportStartedAt);
                    if (fetched < DATA_BATCH_SIZE) {
                        break;
                    }
                }
            }

            if (getFooter(footerHeader)) {
                writer.write("\n");
                writeStreamingCSVHeader(writer, columns, actualDelimiter);
            }
            if (includeCustomFooter(includeCustomFooter)) {
                writer.write("\n");
                writer.write(footerDecorator);
                writer.write("\n");
            }
            writer.flush();
            if (writer.checkError()) {
                throw new IOException("Failed while writing CSV export: " + outputFile);
            }
            completed = true;
        } finally {
            if (!completed && outputFile.exists() && !outputFile.delete()) {
                LogUtil.warn(getClassName(), "Unable to delete incomplete export file: " + outputFile);
            }
            if (!completed) {
                LogUtil.info(getClassName(), getExportFailureMessage("CSV", processedRows, totalRows, exportStartedAt));
            }
        }
        LogUtil.info(getClassName(), getExportCompletionMessage("CSV", processedRows, totalRows, outputFile, exportStartedAt));
        return outputFile;
    }

    private static void writeStreamingCSVHeader(PrintWriter writer, List<DataListColumn> columns, String delimiter) {
        for (int i = 0; i < columns.size(); i++) {
            if (i > 0) {
                writer.write(delimiter);
            }
            writeEscapedCSVValue(writer, columns.get(i).getLabel(), delimiter);
        }
    }

    private static void writeStreamingCSVRow(PrintWriter writer, DataList dataList, Object row, List<DataListColumn> columns, String delimiter, String exportEncrypt) {
        writer.write("\r\n");
        for (int i = 0; i < columns.size(); i++) {
            if (i > 0) {
                writer.write(delimiter);
            }
            String value = getStreamingFormattedValue(dataList, row, columns.get(i), null, exportEncrypt);
            writeEscapedCSVValue(writer, value, delimiter);
        }
        // Do not flush per row. The buffered stream is flushed on close.
    }

    private static void writeEscapedCSVValue(PrintWriter writer, String value, String delimiter) {
        String safeValue = value == null ? "" : value;
        boolean quote = safeValue.contains(delimiter) || safeValue.indexOf('"') >= 0 || safeValue.indexOf('\r') >= 0 || safeValue.indexOf('\n') >= 0;
        if (quote) {
            writer.write('"');
            writer.write(safeValue.replace("\"", "\"\""));
            writer.write('"');
        } else {
            writer.write(safeValue);
        }
    }

    public static void streamCSVFileToResponse(HttpServletResponse response, File csvFile, String filename) throws IOException {
        String name = URLEncoder.encode(filename, "UTF8").replaceAll("\\+", "%20");
        response.setHeader("Content-Disposition", "attachment; filename=" + name + "; filename*=UTF-8''" + name);
        response.setHeader("Content-Length", Long.toString(csvFile.length()));
        response.setContentType("text/csv; charset=UTF-8");
        try (InputStream in = new BufferedInputStream(new java.io.FileInputStream(csvFile), FILE_COPY_BUFFER_SIZE); OutputStream out = new BufferedOutputStream(response.getOutputStream(), FILE_COPY_BUFFER_SIZE)) {
            byte[] buffer = new byte[FILE_COPY_BUFFER_SIZE];
            int length;
            while ((length = in.read(buffer)) != -1) {
                out.write(buffer, 0, length);
            }
        }
    }

    // ---------------------------------------------------------------------
    // New bounded-memory Excel export path
    // ---------------------------------------------------------------------

    /**
     * Generates an XLSX file without retaining the complete workbook or the
     * complete datalist result in JVM memory.
     *
     * The original getExcel(...) method below is intentionally preserved for
     * comparison/review. New call sites should use this method.
     *
     * @param selectedRows rows already filtered for a selected-row export;
     *                     pass null when exporting all rows so this method can
     *                     fetch the datalist in batches
     * @param makeUnique   true for user-configured file paths, false when the
     *                     caller already supplied a unique background-job path
     */
    public static File generateStreamingExcelFile(DataList dataList, DataListCollection selectedRows, String[] rowKeys, File requestedFile, boolean makeUnique, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator, String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportImages, String exportEncrypt, String exportNumeric, Object[] gridColumns) throws IOException {

        long exportStartedAt = System.currentTimeMillis();
        int totalRows = getExpectedExportRows(dataList, selectedRows, rowKeys, downloadAllWhenNoneSelected);
        File outputFile = makeUnique ? getUniqueFile(requestedFile.getPath()) : requestedFile;
        File parent = outputFile.getParentFile();
        if (parent != null && !parent.isDirectory() && !parent.mkdirs()) {
            throw new IOException("Unable to create export directory: " + parent);
        }

        SXSSFWorkbook workbook = new SXSSFWorkbook(SXSSF_ROW_WINDOW);
        // Compress POI's XML temporary files to reduce disk usage for 1M+ rows.
        workbook.setCompressTempFiles(true);
        LogUtil.info(getClassName(), getExportStartMessage("Excel", totalRows, outputFile));

        boolean completed = false;
        long processedRows = 0;
        try {
            StreamingExcelContext context = new StreamingExcelContext(workbook, dataList, headerDecorator, includeCustomHeader, exportImages, exportEncrypt, exportNumeric, gridColumns);

            if (rowKeys != null && rowKeys.length > 0) {
                // Selected-row exports normally contain a relatively small set.
                // A HashSet avoids the legacy rows x selectedKeys nested loop.
                Set<String> selectedKeySet = new HashSet<>(Arrays.asList(rowKeys));
                appendSelectedRows(context, selectedRows, selectedKeySet);
                processedRows = context.getExportedRowCount();
                logExportBatch("Excel", 1, selectedRows != null ? selectedRows.size() : 0, processedRows, totalRows, exportStartedAt);
            } else if ("true".equals(downloadAllWhenNoneSelected)) {
                appendAllRowsInBatches(context, dataList, totalRows, exportStartedAt);
                processedRows = context.getExportedRowCount();
            }

            context.appendFooter(footerHeader, footerDecorator, includeCustomFooter);

            // SXSSFWorkbook has already flushed old rows to disk. This final
            // write assembles the OOXML package directly into the target file.
            LogUtil.info(getClassName(), "TEMP PERF - Excel row processing completed; finalizing workbook: processedRows=" + processedRows + ", elapsedMs=" + (System.currentTimeMillis() - exportStartedAt) + ", usedHeapMB=" + getUsedHeapMB());
            try (OutputStream out = new BufferedOutputStream(new FileOutputStream(outputFile), FILE_COPY_BUFFER_SIZE)) {
                workbook.write(out);
            }
            completed = true;
        } finally {
            try {
                workbook.close();
            } finally {
                // dispose() is required to remove SXSSF worksheet temp files.
                workbook.dispose();
            }
            if (!completed && outputFile.exists() && !outputFile.delete()) {
                LogUtil.warn(getClassName(), "Unable to delete incomplete export file: " + outputFile);
            }
            if (!completed) {
                LogUtil.info(getClassName(), getExportFailureMessage("Excel", processedRows, totalRows, exportStartedAt));
            }
        }
        LogUtil.info(getClassName(), getExportCompletionMessage("Excel", processedRows, totalRows, outputFile, exportStartedAt));
        return outputFile;
    }

    private static void appendAllRowsInBatches(StreamingExcelContext context, DataList dataList, int totalRows, long exportStartedAt) {
        int start = 0;
        int batchNumber = 0;
        while (true) {
            DataListCollection batch = dataList.getRows(DATA_BATCH_SIZE, start);
            if (batch == null || batch.isEmpty()) {
                break;
            }
            for (int i = 0; i < batch.size(); i++) {
                context.appendRow(batch, i);
            }
            int fetched = batch.size();
            batchNumber++;
            start += fetched;
            logExportBatch("Excel", batchNumber, fetched, context.getExportedRowCount(), totalRows, exportStartedAt);
            // A short final batch proves that there are no more records and
            // avoids one additional database query.
            if (fetched < DATA_BATCH_SIZE) {
                break;
            }
        }
    }

    private static void appendSelectedRows(StreamingExcelContext context, DataListCollection rows, Set<String> selectedKeys) {
        if (rows == null) {
            return;
        }
        for (int i = 0; i < rows.size(); i++) {
            if (selectedKeys.contains(findRowKey(rows, i))) {
                context.appendRow(rows, i);
            }
        }
    }

    /**
     * Streams an already generated file to the servlet response without the
     * legacy ByteArrayOutputStream/toByteArray full-file memory copies.
     */
    public static void streamExcelFileToResponse(HttpServletResponse response, File excelFile, String filename) throws IOException {
        String name = URLEncoder.encode(filename, "UTF8").replaceAll("\\+", "%20");
        response.setHeader("Content-Disposition", "attachment; filename=" + name + "; filename*=UTF-8''" + name);
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        // The plugin still targets the Servlet 2.4 API, which does not expose
        // setContentLengthLong(). Setting the header supports files over 2 GB.
        response.setHeader("Content-Length", Long.toString(excelFile.length()));

        try (InputStream in = new BufferedInputStream(new java.io.FileInputStream(excelFile), FILE_COPY_BUFFER_SIZE); OutputStream out = new BufferedOutputStream(response.getOutputStream(), FILE_COPY_BUFFER_SIZE)) {
            byte[] buffer = new byte[FILE_COPY_BUFFER_SIZE];
            int length;
            while ((length = in.read(buffer)) != -1) {
                out.write(buffer, 0, length);
            }
        }
    }

    /** Makes the existing form-storage operation reusable by the new file path. */
    public static void storeGeneratedFileToForm(File generatedFile, String formDefId, String fileFieldId) {
        storeGeneratedFile(generatedFile, formDefId, fileFieldId);
    }

    // Temporary performance logging helpers for the large-dataset test.
    private static int getExpectedExportRows(DataList dataList, DataListCollection selectedRows, String[] rowKeys, String downloadAllWhenNoneSelected) {
        if (rowKeys != null && rowKeys.length > 0) {
            return rowKeys.length;
        }
        if ("true".equals(downloadAllWhenNoneSelected)) {
            return dataList.getTotal();
        }
        return selectedRows != null ? selectedRows.size() : 0;
    }

    private static String getExportStartMessage(String exportType, int totalRows, File outputFile) {
        return "TEMP PERF - " + exportType + " export started: totalRows=" + totalRows + ", dataBatchSize=" + DATA_BATCH_SIZE + ", sxssfRowWindow=" + SXSSF_ROW_WINDOW + ", usedHeapMB=" + getUsedHeapMB() + ", maxHeapMB=" + getMaxHeapMB() + ", output=" + outputFile.getPath();
    }

    private static void logExportBatch(String exportType, int batchNumber, int fetchedRows, long processedRows, int totalRows, long exportStartedAt) {
        long percentage = totalRows > 0 ? Math.min(100, processedRows * 100 / totalRows) : 0;
        LogUtil.info(getClassName(), "TEMP PERF - " + exportType + " batch processed: batch=" + batchNumber + ", fetchedRows=" + fetchedRows + ", processedRows=" + processedRows + ", totalRows=" + totalRows + ", progress=" + percentage + "%, elapsedMs=" + (System.currentTimeMillis() - exportStartedAt) + ", usedHeapMB=" + getUsedHeapMB());
    }

    private static String getExportCompletionMessage(String exportType, long processedRows, int totalRows, File outputFile, long exportStartedAt) {
        return "TEMP PERF - " + exportType + " export completed: processedRows=" + processedRows + ", totalRows=" + totalRows + ", elapsedMs=" + (System.currentTimeMillis() - exportStartedAt) + ", fileSizeBytes=" + outputFile.length() + ", usedHeapMB=" + getUsedHeapMB() + ", output=" + outputFile.getPath();
    }

    private static String getExportFailureMessage(String exportType, long processedRows, int totalRows, long exportStartedAt) {
        return "TEMP PERF - " + exportType + " export failed: processedRows=" + processedRows + ", totalRows=" + totalRows + ", elapsedMs=" + (System.currentTimeMillis() - exportStartedAt) + ", usedHeapMB=" + getUsedHeapMB();
    }

    private static long getUsedHeapMB() {
        Runtime runtime = Runtime.getRuntime();
        return (runtime.totalMemory() - runtime.freeMemory()) / (1024 * 1024);
    }

    private static long getMaxHeapMB() {
        return Runtime.getRuntime().maxMemory() / (1024 * 1024);
    }

    private static final class StreamingExcelContext {
        private final SXSSFWorkbook workbook;
        private final DataList dataList;
        private final List<DataListColumn> columns;
        private final String[] headers;
        private final String headerDecorator;
        private final boolean includeHeaderDecorator;
        private final String exportImages;
        private final String exportEncrypt;
        private final Set<String> numericColumns;
        private final CellStyle numericStyle;
        private final AppDefinition appDef;
        private Sheet sheet;
        private StreamingImageSupport imageSupport;
        private int sheetNumber;
        private int rowNumber;
        private long exportedRowCount;

        private StreamingExcelContext(SXSSFWorkbook workbook, DataList dataList, String headerDecorator, String includeCustomHeader, String exportImages, String exportEncrypt, String exportNumeric, Object[] gridColumns) {
            this.workbook = workbook;
            this.dataList = dataList;
            this.columns = getExportColumns(dataList);
            this.headers = new String[columns.size()];
            for (int i = 0; i < columns.size(); i++) {
                headers[i] = columns.get(i).getLabel();
            }
            this.headerDecorator = headerDecorator;
            this.includeHeaderDecorator = includeCustomHeader(includeCustomHeader);
            this.exportImages = exportImages;
            this.exportEncrypt = exportEncrypt;
            this.numericColumns = getNumericColumns(exportNumeric, gridColumns);
            this.numericStyle = createStreamingNumericStyle(workbook);
            this.appDef = AppUtil.getCurrentAppDefinition();
            createSheet();
        }

        private void createSheet() {
            sheetNumber++;
            sheet = workbook.createSheet(sheetNumber == 1 ? "Report" : "Report " + sheetNumber);
            imageSupport = "true".equals(exportImages) ? new StreamingImageSupport(workbook, sheet, appDef) : null;
            rowNumber = 0;

            if (includeHeaderDecorator) {
                Row titleRow = sheet.createRow(rowNumber++);
                Cell titleCell = titleRow.createCell(0);
                titleCell.setCellValue(headerDecorator);
                int lineCount = headerDecorator.split("\\r\\n|\\r|\\n").length;
                titleRow.setHeightInPoints(lineCount * sheet.getDefaultRowHeightInPoints());
                if (headers.length >= 2) {
                    sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, headers.length - 1));
                }
            }

            Row headerRow = sheet.createRow(rowNumber++);
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }
        }

        private void appendRow(DataListCollection sourceRows, int sourceIndex) {
            if (rowNumber >= XLSX_MAX_ROWS_PER_SHEET) {
                // XLSX worksheets are limited to 1,048,576 rows. Continue in a
                // new sheet and repeat the configured/header rows.
                createSheet();
            }

            Row excelRow = sheet.createRow(rowNumber);
            Object sourceRow = getRow(sourceRows, sourceIndex);
            int columnNumber = 0;
            for (DataListColumn column : columns) {
                String value = getStreamingFormattedValue(dataList, sourceRow, column, exportImages, exportEncrypt);
                if (value.startsWith("IMAGE:") || value.startsWith("FILE:")) {
                    // Image data is still retained by POI at workbook scope;
                    // this feature requires separate high-volume testing.
                    imageSupport.append(excelRow, rowNumber, sourceRow, value, columnNumber++);
                } else {
                    Cell cell = excelRow.createCell(columnNumber++);
                    if (numericColumns.contains(column.getName())
                            && NumberUtils.isParsable(value)) {
                        cell.setCellStyle(numericStyle);
                        cell.setCellValue(Double.parseDouble(value));
                    } else {
                        cell.setCellValue(value);
                    }
                }
            }
            rowNumber++;
            exportedRowCount++;
        }

        private long getExportedRowCount() {
            return exportedRowCount;
        }

        private void appendFooter(String footerHeader, String footerDecorator, String includeCustomFooter) {
            int footerRows = (getFooter(footerHeader) ? 1 : 0) + (includeCustomFooter(includeCustomFooter) ? 1 : 0);
            if (footerRows > 0 && rowNumber + footerRows > XLSX_MAX_ROWS_PER_SHEET) {
                createSheet();
            }
            if (getFooter(footerHeader)) {
                Row footerHeaderRow = sheet.createRow(rowNumber++);
                for (int i = 0; i < headers.length; i++) {
                    footerHeaderRow.createCell(i).setCellValue(headers[i]);
                }
            }
            if (includeCustomFooter(includeCustomFooter)) {
                Row footerRow = sheet.createRow(rowNumber++);
                footerRow.createCell(0).setCellValue(footerDecorator);
                if (headers.length >= 2) {
                    sheet.addMergedRegion(new CellRangeAddress(footerRow.getRowNum(), footerRow.getRowNum(), 0, headers.length - 1));
                }
            }
        }
    }

    /**
     * Reuses services, Tika, POI helpers and the drawing patriarch instead of
     * recreating them for every image cell as the legacy method does.
     */
    private static final class StreamingImageSupport {
        private final Workbook workbook;
        private final Sheet sheet;
        private final AppDefinition appDef;
        private final AppService appService;
        private final Tika tika = new Tika();
        private final CreationHelper creationHelper;
        private final Drawing<?> drawing;
        private final Map<String, String> tableNames = new HashMap<>();

        private StreamingImageSupport(Workbook workbook, Sheet sheet, AppDefinition appDef) {
            this.workbook = workbook;
            this.sheet = sheet;
            this.appDef = appDef;
            ApplicationContext context = AppUtil.getApplicationContext();
            this.appService = (AppService) context.getBean("appService");
            this.creationHelper = workbook.getCreationHelper();
            this.drawing = sheet.createDrawingPatriarch();
        }

        private void append(Row excelRow, int rowNumber, Object sourceRow, String encodedValue, int columnNumber) {
            excelRow.createCell(columnNumber);
            try {
                String[] pieces;
                String formDefId;
                String fileName;
                boolean knownImage;
                if (encodedValue.startsWith("IMAGE:")) {
                    pieces = encodedValue.split(":", 4);
                    if (pieces.length < 4) {
                        return;
                    }
                    formDefId = pieces[1];
                    fileName = pieces[3];
                    knownImage = true;
                } else {
                    pieces = encodedValue.split(":", 3);
                    if (pieces.length < 3) {
                        return;
                    }
                    formDefId = pieces[1];
                    fileName = pieces[2];
                    knownImage = false;
                }

                String tableName = tableNames.get(formDefId);
                if (tableName == null) {
                    tableName = appService.getFormTableName(appDef, formDefId);
                    tableNames.put(formDefId, tableName);
                }
                File file = FileUtil.getFile(fileName, tableName, findRowKey(sourceRow));
                if (file == null || !file.exists()) {
                    return;
                }
                String mimeType = knownImage ? "image/known" : tika.detect(file);
                if (mimeType == null || !mimeType.startsWith("image/")) {
                    return;
                }

                byte[] imageBytes = Files.readAllBytes(file.toPath());
                int pictureIndex = workbook.addPicture(imageBytes, getPictureType(fileName));
                ClientAnchor anchor = creationHelper.createClientAnchor();
                sheet.setColumnWidth(columnNumber, 1500);
                excelRow.setHeightInPoints(40);
                anchor.setCol1(columnNumber);
                anchor.setRow1(rowNumber);
                anchor.setCol2(columnNumber + 1);
                anchor.setRow2(rowNumber + 1);
                drawing.createPicture(anchor, pictureIndex).resize(1.0, 1.0);
            } catch (IOException ex) {
                LogUtil.error(getClassName(), ex, ex.getMessage());
            }
        }
    }

    private static String findRowKey(Object row) {
        Object idValue = null;
        if (row instanceof Map) {
            idValue = ((Map) row).get("id");
        } else if (row instanceof FormRow) {
            idValue = ((FormRow) row).get("id");
        }
        return idValue != null ? idValue.toString() : null;
    }

    private static List<DataListColumn> getExportColumns(DataList dataList) {
        List<DataListColumn> exportColumns = new ArrayList<>();
        for (DataListColumn column : dataList.getColumns()) {
            String excludeExport = column.getPropertyString("exclude_export");
            String includeExport = column.getPropertyString("include_export");
            boolean hidden = column.isHidden();
            if ((hidden && "true".equalsIgnoreCase(includeExport))
                    || (!hidden && !"true".equalsIgnoreCase(excludeExport))) {
                exportColumns.add(column);
            }
        }
        return exportColumns;
    }

    private static Set<String> getNumericColumns(String exportNumeric, Object[] gridColumns) {
        if (!"true".equals(exportNumeric) || gridColumns == null) {
            return Collections.emptySet();
        }
        Set<String> numericColumns = new HashSet<>();
        for (Object gridColumn : gridColumns) {
            if (gridColumn instanceof Map) {
                Object field = ((Map) gridColumn).get("field");
                if (field != null) {
                    numericColumns.add(field.toString());
                }
            }
        }
        return numericColumns;
    }

    private static String getStreamingFormattedValue(DataList dataList, Object row, DataListColumn column, String exportImages, String exportEncrypt) {
        String name = column.getName();
        try {
            Object valueObj = null;
            if (column instanceof DataListDisplayColumnProxy) {
                Object displayColumn = ((DataListDisplayColumnProxy) column)
                        .getDisplayColumn();
                if (displayColumn instanceof BeanShellColumn) {
                    valueObj = ((BeanShellColumn) displayColumn)
                            .getRowValue(row, 0);
                }
            }
            if (valueObj == null) {
                valueObj = DataListService.evaluateColumnValueFromRow(row, name);
            }
            String value = valueObj != null ? valueObj.toString() : "";
            Collection<DataListColumnFormat> formats = column.getFormats();

            if ("true".equals(exportImages)
                    && formats != null && !formats.isEmpty()) {
                DataListColumnFormat firstFormat = formats.iterator().next();
                if (firstFormat != null) {
                    String formatterClassName = firstFormat.getClassName();
                    String formDefId = (String) firstFormat.getProperty("formDefId");
                    if ("org.joget.apps.datalist.lib.ImageFormatter"
                            .equals(formatterClassName)) {
                        String imageSrc = (String) firstFormat.getProperty("imageSrc");
                        if ("form".equals(imageSrc) && !value.isEmpty()) {
                            return "IMAGE:" + formDefId + ":" + imageSrc
                                    + ":" + value;
                        }
                    } else if ("org.joget.tutorial.FileLinkDatalistFormatter"
                            .equals(formatterClassName)
                            && formDefId != null && !formDefId.isEmpty()
                            && !value.isEmpty()) {
                        return "FILE:" + formDefId + ":" + value;
                    }
                }
            }

            if (!"true".equals(exportEncrypt)) {
                value = SecurityUtil.decrypt(value);
            }
            // Preserve legacy behavior: only the first non-null formatter is
            // applied before returning the exported value.
            if (formats != null) {
                for (DataListColumnFormat format : formats) {
                    if (format != null) {
                        value = format.format(dataList, column, row, value);
                        break;
                    }
                }
            }
            return value == null ? ""
                    : HTML_TAG_PATTERN.matcher(value).replaceAll("");
        } catch (Exception ex) {
            LogUtil.error(getClassName(), ex, "Error processing column : " + name);
            return "";
        }
    }

    /**
     * Legacy in-memory implementation retained for reference and backwards
     * compatibility. New plugin call sites use generateStreamingExcelFile().
     */
    public static Workbook getExcel(DataList dataList, DataListCollection rows, String[] rowKeys, boolean background, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator, String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportImages, String exportEncrypt, String exportNumeric, Object[] gridColumns) {
        HashMap<String, StringBuilder> sb = getLabelAndKey(dataList);
        StringBuilder keySB = sb.get("key");
        StringBuilder headerSB = sb.get("header");
        int counter = 0;
        int rowCounter = 0;

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Report");
        Row headerRow = sheet.createRow(rowCounter);
        String[] res = keySB.toString().split(",", 0);
        String[] header = headerSB.toString().split(",", 0);
        duplicates.setMap(findDuplicate(res));

        if (includeCustomHeader(includeCustomHeader)) {
            Cell titleCell = headerRow.createCell(0);
            String headerString = headerDecorator;
            titleCell.setCellValue(headerString);
            int getNewLine = headerString.split("\r\n|\r|\n").length;
            headerRow.setHeightInPoints((getNewLine * sheet.getDefaultRowHeightInPoints()));

            if (header.length >= 2) {
                sheet.autoSizeColumn(2);
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, header.length - 1));
            }
            rowCounter += 1;
        }

        Row headerColumnRow = sheet.createRow(rowCounter);
        counter = 0;
        for (String value : header) {
            Cell headerCell = headerColumnRow.createCell(counter);
            headerCell.setCellValue(value);
            counter += 1;
        }

        rowCounter += 1;
        counter = 0;

        AppDefinition currentAppDef = AppUtil.getCurrentAppDefinition();

        if (rowKeys != null && rowKeys.length > 0) {
            if (!dataList.isUseSession()) {
                // rows already filtered by primary key IN (...); skip "id" matching
                for (int x = 0; x < rows.size(); x++) {
                    printExcel(currentAppDef, sheet, rowCounter, counter, rows, x, res, dataList, exportImages, exportEncrypt, exportNumeric, gridColumns);
                    counter += 1;
                    rowCounter += 1;
                }
            } else {
                for (int x = 0; x < rows.size(); x++) {
                    //compare with all the rowkeys that have been selected
                    for (int y = 0; y < rowKeys.length; y++) {
                        boolean boolInstance = rows.get(x) instanceof HashMap;
                        boolean foundRowKey = foundRowKey(boolInstance, rows, x, rowKeys[y]);

                        if (!foundRowKey) {
                            continue;
                        }
                        printExcel(currentAppDef, sheet, rowCounter, counter, rows, x, res, dataList, exportImages, exportEncrypt, exportNumeric, gridColumns);
                        counter += 1;
                        rowCounter += 1;
                        break;
                    }
                }
            }

        } else if (downloadAllWhenNoneSelected.equals("true")) {
            for (int x = 0; x < rows.size(); x++) {
                printExcel(currentAppDef, sheet, rowCounter, counter, rows, x, res, dataList, exportImages, exportEncrypt, exportNumeric, gridColumns);
                counter += 1;
                rowCounter += 1;
            }
        }

        if (getFooter(footerHeader)) {
            int z = 0;
            Row dataRow = sheet.createRow(rowCounter);
            for (String myStr : header) {
                Cell footerCell = dataRow.createCell(z);
                footerCell.setCellValue(myStr);
                z += 1;
            }
            rowCounter += 1;
        }

        if (includeCustomFooter(includeCustomFooter)) {
            Row footerColumnRow = sheet.createRow(rowCounter);
            Cell titleCell = footerColumnRow.createCell(0);
            titleCell.setCellValue(footerDecorator);

            if (header.length >= 2) {
                sheet.addMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, header.length - 1));
            }
        }

        sheet.setForceFormulaRecalculation(true);
        return workbook;

    }

    public static boolean checkCompletionFlag(String path) {
        File flagFile = new File(path + ".completed");
        return flagFile.exists();
    }

    public static void downloadExcel(HttpServletRequest request, HttpServletResponse response, DataList dataList, DataListCollection dataListRows, String[] rowKeys, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator, String renameFile, String fileName, String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportImages, String exportEncrypt, String exportNumeric, Object[] gridColumns) throws ServletException, IOException {
        Workbook workbook = getExcel(dataList, dataListRows, rowKeys, false, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt, exportNumeric, gridColumns);
        String filename = renameFile.equalsIgnoreCase("true") ? fileName + ".xlsx" : "report.xlsx";
        writeResponseExcel(request, response, workbook, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\n");
    }

    protected static HashMap<String, Integer> findDuplicate(String[] keySB) {
        Set<String> seen = new HashSet<>();
        HashMap<String, Integer> duplicates = new HashMap<>();

        for (String str : keySB) {
            if (seen.contains(str)) {
                if (!duplicates.containsKey(str)) {
                    duplicates.put(str, 1);
                } else {
                    duplicates.put(str, duplicates.get(str) + 1);
                }
            } else {
                seen.add(str);
            }
        }
        return duplicates;
    }

    protected static void writeResponseExcel(HttpServletRequest request, HttpServletResponse response, Workbook workbook, String filename, String contentType) throws IOException, ServletException {
        OutputStream out = response.getOutputStream();
        try {
            String name = URLEncoder.encode(filename, "UTF8").replaceAll("\\+", "%20");
            response.setHeader("Content-Disposition", "attachment; filename=" + name + "; filename*=UTF-8''" + name);
            response.setContentType(contentType + "; charset=UTF-8");

            ByteArrayOutputStream ms = new ByteArrayOutputStream();
            workbook.write(ms);

            byte bytes[] = ms.toByteArray();
            if (bytes.length > 0) {
                response.setContentLength(bytes.length);
                out.write(bytes);
            }

        } finally {
            out.flush();
            out.close();
            request.getRequestDispatcher(filename).forward(request, response);
        }
    }

    protected static String getBinderFormattedValue(DataList dataList, Object o, String name,
            String exportImages, String exportEncrypt) {

        DataListColumn[] columns = dataList.getColumns();
        int skip = duplicates.getSkipCount(name);
        for (DataListColumn c : columns) {
            if (c.getName().equalsIgnoreCase(name)) {

                if ("true".equals(exportImages)) {
                    Collection<DataListColumnFormat> formatsList = c.getFormats();
                    if (formatsList != null && !formatsList.isEmpty()) {
                        DataListColumnFormat firstFormat = formatsList.iterator().next();

                        if (firstFormat != null) {
                            String formatterClassName = firstFormat.getClassName();
                            String filename = DataListService.evaluateColumnValueFromRow(o, name).toString();
                            if ("org.joget.apps.datalist.lib.ImageFormatter".equals(formatterClassName)) {

                                String formDefId = (String) firstFormat.getProperty("formDefId");
                                String imageSrc = (String) firstFormat.getProperty("imageSrc");

                                if ("form".equals(imageSrc) && filename != null && !filename.isEmpty()) {
                                    return "IMAGE:" + formDefId + ":" + imageSrc + ":" + filename;
                                }
                            } else if ("org.joget.tutorial.FileLinkDatalistFormatter".equals(formatterClassName)) {
                                // file upload field
                                String formDefId = (String) firstFormat.getProperty("formDefId");

                                if (formDefId != null && !formDefId.isEmpty()
                                        && filename != null && !filename.isEmpty()) {

                                    return "FILE:" + formDefId + ":" + filename;
                                }
                            }
                        }
                    }
                }

                if (duplicates.checkKey(name)) {
                    if (duplicates.skipCountLessThenDuplicate(name)) {
                        duplicates.addSkipCount(name);
                    }
                    if (skip != 0) {
                        skip -= 1;
                        continue;
                    }
                }

                try {

                    Object valueObj = null;

                    if (c instanceof DataListDisplayColumnProxy) {

                        DataListDisplayColumnProxy proxy
                                = (DataListDisplayColumnProxy) c;

                        Object displayColumn = proxy.getDisplayColumn();

                        if (displayColumn instanceof BeanShellColumn) {

                            valueObj
                                    = ((BeanShellColumn) displayColumn)
                                            .getRowValue(o, 0);
                        }
                    }

                    if (valueObj == null) {
                        valueObj = DataListService.evaluateColumnValueFromRow(o, name);
                    }

                    String value = valueObj != null ? valueObj.toString() : "";

                    if (!"true".equals(exportEncrypt)) {
                        value = SecurityUtil.decrypt(value);
                    }

                    Collection<DataListColumnFormat> formats = c.getFormats();

                    if (formats != null) {
                        for (DataListColumnFormat f : formats) {

                            if (f != null) {

                                value = f.format(dataList, c, o, value);

                                String stripHTML
                                        = value.replaceAll("<[^>]*>", "");

                                return stripHTML;
                            } else {
                                return value;
                            }
                        }
                    } else {
                        return value;
                    }

                } catch (Exception ex) {
                    LogUtil.error(getClassName(), ex,
                            "Error processing column : " + name);
                }
            }
        }

        return "";
    }

    protected static HashMap<String, StringBuilder> getLabelAndKey(DataList dataList) {
        HashMap<String, StringBuilder> sb = new HashMap<>();
        StringBuilder headerSB = new StringBuilder();
        StringBuilder keySB = new StringBuilder();

        for (DataListColumn column : dataList.getColumns()) {
            String header = column.getLabel();
            String key = column.getName();

            String excludeExport = column.getPropertyString("exclude_export");
            String includeExport = column.getPropertyString("include_export");
            boolean hidden = column.isHidden();

            if ((hidden && "true".equalsIgnoreCase(includeExport)) || (!hidden && !"true".equalsIgnoreCase(excludeExport))) {
                headerSB.append(header).append(",");
                keySB.append(key).append(",");
            }
        }
        headerSB.setLength(headerSB.length() - 1);
        keySB.setLength(keySB.length() - 1);

        sb.put("header", headerSB);
        sb.put("key", keySB);
        return sb;
    }

    protected static void printExcel(AppDefinition currentAppDef, Sheet sheet, int rowCounter, int counter, DataListCollection rows, int x, String[] res, DataList dataList, String exportImages, String exportEncrypt, String exportNumeric, Object[] gridColumns) {
        Row dataRow = sheet.createRow(rowCounter);
        Object row = getRow(rows, x);
        int z = 0;

        // number format for excel
        Workbook workbook = sheet.getWorkbook();
        CellStyle numberStyle = getNumericStyle(sheet.getWorkbook());

        // get numeric column from config
        Set<String> numericColumns = new HashSet<>();

        if ("true".equals(exportNumeric)) {
            if (gridColumns != null && gridColumns.length > 0) {
                for (Object o : gridColumns) {
                    Map mapping = (HashMap) o;
                    String columnField = mapping.get("field").toString();
                    numericColumns.add(columnField);
                }
            }
        }

        for (String myStr : res) {
            String value = getBinderFormattedValue(dataList, row, myStr, exportImages, exportEncrypt);

            // Check if value contains multiple images (if they're separated by some delimiter)
            if (value.startsWith("IMAGE:") || value.startsWith("FILE:")) {
                // Process image/file - this might consume multiple columns if there are multiple images
                z = processImageOrFile(currentAppDef, sheet, dataRow, rowCounter, rows, x, value, z);
            } else {
                Cell dataRowCell = dataRow.createCell(z);

                // check if this column is numeric
                if ("true".equals(exportNumeric) && numericColumns.contains(myStr) && NumberUtils.isParsable(value)) {
                    double numericValue = Double.parseDouble(value);
                    dataRowCell.setCellStyle(numberStyle);
                    dataRowCell.setCellValue(numericValue);
                } else {
                    dataRowCell.setCellValue(value);
                }

                z++;
            }
        }
    }

    private static int processImageOrFile(AppDefinition currentAppDef, Sheet sheet, Row dataRow, int rowCounter, DataListCollection rows, int x, String value, int startColumn) {
        int currentColumn = startColumn;

        // If your value can contain multiple images separated by some delimiter, split them here
        // For now, assuming single image per field
        String[] imageValues = {value}; // Modify this if you have multiple images in one field

        for (String imageValue : imageValues) {
            Cell dataRowCell = dataRow.createCell(currentColumn);

            if (imageValue.startsWith("IMAGE:")) {
                try {
                    // image found
                    String rowId = findRowKey(rows, x);

                    ApplicationContext ac = AppUtil.getApplicationContext();
                    AppService appService = (AppService) ac.getBean("appService");
                    String[] pieces = imageValue.split(":");
                    String formDefId = pieces[1];
                    String imageSrc = pieces[2];
                    String fileName = pieces[3];

                    String tableName = appService.getFormTableName(currentAppDef, formDefId);
                    File imageFile = FileUtil.getFile(fileName, tableName, rowId);
                    if (imageFile != null && imageFile.exists()) {
                        byte[] imageBytes = Files.readAllBytes(imageFile.toPath());
                        Workbook workbook = sheet.getWorkbook();
                        int pictureType = getPictureType(fileName);
                        int pictureIdx = workbook.addPicture(imageBytes, pictureType);
                        CreationHelper helper = workbook.getCreationHelper();
                        Drawing<?> drawing = sheet.createDrawingPatriarch();
                        ClientAnchor anchor = helper.createClientAnchor();

                        // Create thumbnail-sized cell
                        int thumbnailWidth = 1500;  // Column width units
                        float thumbnailHeight = 40; // Row height in points

                        sheet.setColumnWidth(currentColumn, thumbnailWidth);
                        dataRow.setHeightInPoints(thumbnailHeight);

                        anchor.setCol1(currentColumn);
                        anchor.setRow1(rowCounter);
                        anchor.setCol2(currentColumn + 1);
                        anchor.setRow2(rowCounter + 1);

                        Picture pict = drawing.createPicture(anchor, pictureIdx);

                        // Scale to fit nicely in thumbnail cell
                        pict.resize(1.0, 1.0); // Adjust between 0.4 to 1.0 based on your preference
                    }
                } catch (IOException ex) {
                    LogUtil.error(getClassName(), ex, ex.getMessage());
                }
            } else if (imageValue.startsWith("FILE:")) {
                ApplicationContext ac = AppUtil.getApplicationContext();
                AppService appService = (AppService) ac.getBean("appService");
                String[] pieces = imageValue.split(":");
                String formDefId = pieces[1];
                String fileName = pieces[2];
                String rowId = findRowKey(rows, x);

                String tableName = appService.getFormTableName(currentAppDef, formDefId);
                try {
                    File file = FileUtil.getFile(fileName, tableName, rowId);
                    if (file != null && file.exists()) {
                        Tika tika = new Tika();
                        String mimeType = tika.detect(file);
                        if (mimeType != null && mimeType.startsWith("image/")) {
                            byte[] imageBytes = Files.readAllBytes(file.toPath());
                            Workbook workbook = sheet.getWorkbook();
                            int pictureType = getPictureType(fileName);
                            int pictureIdx = workbook.addPicture(imageBytes, pictureType);
                            CreationHelper helper = workbook.getCreationHelper();
                            Drawing<?> drawing = sheet.createDrawingPatriarch();
                            ClientAnchor anchor = helper.createClientAnchor();

                            // Create thumbnail-sized cell
                            int thumbnailWidth = 1500;  // Column width units
                            float thumbnailHeight = 40; // Row height in points

                            sheet.setColumnWidth(currentColumn, thumbnailWidth);
                            dataRow.setHeightInPoints(thumbnailHeight);

                            anchor.setCol1(currentColumn);
                            anchor.setRow1(rowCounter);
                            anchor.setCol2(currentColumn + 1);
                            anchor.setRow2(rowCounter + 1);

                            Picture pict = drawing.createPicture(anchor, pictureIdx);

                            // Scale to fit nicely in thumbnail cell
                            pict.resize(1.0, 1.0); // Adjust between 0.4 to 1.0 based on your preference
                        }
                    }
                } catch (IOException ex) {
                    LogUtil.error(getClassName(), ex, ex.getMessage());
                }
            }

            currentColumn += 1; // Move to next column for next image
        }

        return currentColumn; // Return the next available column index
    }

    protected static String findRowKey(DataListCollection rows, int x) {
        Object row = rows.get(x);
        Object idValue = null;

        if (row instanceof HashMap) {
            idValue = ((HashMap) row).get("id");
        } else if (row instanceof FormRow) {
            idValue = ((FormRow) row).get("id");
        }

        return (idValue != null) ? idValue.toString() : null;
    }

    private static int getPictureType(String fileName) {
        String extension = fileName.toLowerCase();
        if (extension.endsWith(".png")) {
            return Workbook.PICTURE_TYPE_PNG;
        } else if (extension.endsWith(".jpg") || extension.endsWith(".jpeg")) {
            return Workbook.PICTURE_TYPE_JPEG;
        } else {
            return Workbook.PICTURE_TYPE_PNG; // Default fallback
        }
    }

    protected static boolean isWholeNumber(double value) {
        // Check if the value is a whole number (i.e., has no decimal points)
        return value == Math.floor(value) && !Double.isInfinite(value);
    }

    protected static Object getRow(DataListCollection rows, int x) {
        return rows.get(x);
    }

    protected static boolean foundRowKey(boolean boolInstance, DataListCollection rows, int x, String rowKey) {
        Object idValue = boolInstance
                ? ((HashMap) rows.get(x)).get("id")
                : ((FormRow) rows.get(x)).get("id");
        return idValue != null && idValue.equals(rowKey);
    }

    protected static boolean getFooter(String footerHeader) {
        String footer = footerHeader;
        return footer.equalsIgnoreCase("true");
    }

    protected static boolean includeCustomHeader(String includeCustomHeader) {
        String header = includeCustomHeader;
        return header.equalsIgnoreCase("true");
    }

    protected static boolean includeCustomFooter(String includeCustomFooter) {
        String footer = includeCustomFooter;
        return footer.equalsIgnoreCase("true");
    }

    public static String getClassName() {
        return "DownloadCsvOrExcelUtil";
    }

    protected static File generateCSVOutputFile(String content, String fileName) throws IOException {
        File outFile = getUniqueFile(fileName);

        try (PrintWriter writer = new PrintWriter(new FileWriter(outFile))) {
            writer.write(content);
        }

        return outFile;
    }

    public static File generateExcelOutputFile(Workbook workbook, String fileName) throws IOException {
        File outFile = getUniqueFile(fileName);

        try (FileOutputStream out = new FileOutputStream(outFile)) {
            workbook.write(out);
        }

        workbook.close();

        return outFile;
    }

    protected static File getUniqueFile(String fileName) {
        File file = new File(fileName);

        if (!file.exists()) {
            return file;
        }

        String name = file.getName();
        String parent = file.getParent();
        if (parent == null) {
            parent = ".";
        }

        String baseName;
        String extension = "";

        int dotIndex = name.lastIndexOf('.');
        if (dotIndex > 0 && dotIndex < name.length() - 1) {
            baseName = name.substring(0, dotIndex);
            extension = name.substring(dotIndex); // includes the dot
        } else {
            baseName = name;
        }

        int counter = 1;
        File newFile;
        do {
            String newName = baseName + " (" + counter + ")" + extension;
            newFile = new File(parent, newName);
            counter++;
        } while (newFile.exists());

        return newFile;
    }
}
