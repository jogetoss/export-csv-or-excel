package org.joget.marketplace;

import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListActionDefault;
import org.joget.apps.datalist.model.DataListActionResult;
import org.joget.apps.datalist.model.DataListCollection;
import org.joget.apps.datalist.model.DataListFilterQueryObject;
import org.joget.commons.util.FileManager;
import org.joget.commons.util.LogUtil;
import org.joget.workflow.model.service.WorkflowUserManager;
import org.joget.workflow.util.WorkflowUtil;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;

import org.joget.apps.app.model.AppDefinition;
import org.joget.commons.util.PluginThread;
import org.joget.commons.util.UuidGenerator;
import org.joget.marketplace.util.DownloadCsvOrExcelUtil;
import org.joget.marketplace.util.BackgroundExportStatus;
import org.joget.plugin.base.PluginWebSupport;

public class DownloadCsvOrExcelDatalistAction extends DataListActionDefault implements PluginWebSupport {

    private final static String MESSAGE_PATH = "messages/DownloadCSVOrExcelDatalistAction";

    @Override
    public String getName() {
        return "Download CSV or Excel Datalist Action";
    }

    @Override
    public String getVersion() {
        return Activator.VERSION;
    }

    @Override
    public String getClassName() {
        return getClass().getName();
    }

    @Override
    public String getLabel() {
        //support i18n
        return AppPluginUtil.getMessage("org.joget.DownloadCSVOrExcelDatalistAction.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getDescription() {
        //support i18n
        return AppPluginUtil.getMessage("org.joget.DownloadCSVOrExcelDatalistAction.pluginDesc", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/DownloadCSVOrExcelDatalistAction.json", null, true, MESSAGE_PATH);
    }

    @Override
    public String getLinkLabel() {
        return getPropertyString("label"); //get label from configured properties options
    }

    @Override
    public String getHref() {
        return getPropertyString("href"); //Let system to handle to post to the same page
    }

    @Override
    public String getTarget() {
        String downloadBackgroud = getPropertyString("downloadBackgroud");
        if ("true".equals(downloadBackgroud)) {
            return "_blank";
        }
        return "post";
    }

    @Override
    public String getHrefParam() {
        return getPropertyString("hrefParam");  //Let system to set the parameter to the checkbox name
    }

    @Override
    public String getHrefColumn() {
        String recordIdColumn = getPropertyString("recordIdColumn"); //get column id from configured properties options
        if ("id".equalsIgnoreCase(recordIdColumn) || recordIdColumn.isEmpty()) {
            return getPropertyString("hrefColumn"); //Let system to set the primary key column of the binder
        } else {
            return recordIdColumn;
        }
    }

    @Override
    public String getConfirmation() {
        return getPropertyString("confirmation"); //get confirmation from configured properties options
    }

    public boolean getDownloadAs() {
        String downloadAs = getPropertyString("downloadAs");
        return downloadAs.equalsIgnoreCase("csv");
    }

    @Override
    public DataListActionResult executeAction(final DataList dataList, String[] rowKeys) {
        String renameFile = getPropertyString("renameFile");
        String fileName = getPropertyString("filename");
        String delimiter = getPropertyString("delimiter");
        final int batchSize = DownloadCsvOrExcelUtil.getDataBatchSize(getPropertyString("dataBatchSize"));
        String headerDecorator = getPropertyString("headerDecorator");
        String downloadAllWhenNoneSelected = getPropertyString("downloadAllWhenNoneSelected");
        String footerDecorator = getPropertyString("footerDecorator");
        String includeCustomHeader = getPropertyString("includeCustomHeader");
        String footerHeader = getPropertyString("footerHeader");
        String includeCustomFooter = getPropertyString("includeCustomFooter");
        String formDefId = getPropertyString("formDefId");
        String fileFieldId = getPropertyString("fileFieldId");
        String exportImages = getPropertyString("exportImages");
        String exportEncrypt = getPropertyString("exportEncrypt");
        String exportNumeric = getPropertyString("exportNumeric");
        Object[] selectedNumericColumn = (Object[]) properties.get("selectedNumericColumn");
        final DownloadCsvOrExcelUtil.HeaderStyle headerStyle = DownloadCsvOrExcelUtil.HeaderStyle.of(
                getPropertyString("headerBackgroundColor"), getPropertyString("headerFontColor"),
                getPropertyString("headerBold"), getPropertyString("headerItalic"),
                getPropertyString("headerFontName"), getPropertyString("headerFontSize"),
                getPropertyString("headerAlignment"));

        // only allow POST
        DataListActionResult result = new DataListActionResult();
        result.setType(DataListActionResult.TYPE_REDIRECT);
        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }
        // check for submited rows
        if ((rowKeys != null && rowKeys.length > 0) || getProperty("downloadAllWhenNoneSelected").equals("true")) {
            try {
                //get the HTTP Response
                HttpServletResponse response = WorkflowUtil.getHttpServletResponse();
                boolean storeToForm = "true".equalsIgnoreCase(getPropertyString("storeToForm"));

                if (getDownloadAs()) {
                    /*
                     * The original downloadCSV/storeCSVToForm methods remain
                     * available in the utility class. This new path pages the
                     * data and writes through a buffered temporary file.
                     */
                    DataListCollection selectedRows = getSelectedRowsForExport(dataList, rowKeys);
                    String csvFileName = renameFile.equalsIgnoreCase("true") ? fileName + ".csv" : "report.csv";
                    File tempFolder = new File(FileManager.getBaseDirectory(), UuidGenerator.getInstance().getUuid());
                    File csvFile = new File(tempFolder, csvFileName);
                    try {
                        DownloadCsvOrExcelUtil.generateStreamingCSVFile(dataList, selectedRows, rowKeys, csvFile, false, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportEncrypt, batchSize);
                        if (storeToForm) {
                            DownloadCsvOrExcelUtil.storeGeneratedFileToForm(csvFile, formDefId, fileFieldId);
                        } else {
                            DownloadCsvOrExcelUtil.streamCSVFileToResponse(response, csvFile, csvFileName);
                        }
                    } finally {
                        deleteTemporaryExport(csvFile, tempFolder);
                    }
                } else {
                    String downloadBackgroud = getPropertyString("downloadBackgroud");
                    if ("true".equalsIgnoreCase(downloadBackgroud)) {
                        String uniqueId = UuidGenerator.getInstance().getUuid();
                        String excelFileName = getPropertyString("renameFile").equalsIgnoreCase("true") ? getPropertyString("filename") + ".xlsx" : "report.xlsx";
                        File excelFolder = new File(FileManager.getBaseDirectory(), uniqueId);
                        if (!excelFolder.isDirectory()) {
                            //create directories if not exist
                            new File(FileManager.getBaseDirectory(), uniqueId).mkdirs();
                        }
                        final BackgroundExportStatus progress = new BackgroundExportStatus(excelFolder, storeToForm);
                        progress.update("preparing", 0, 0);
                        AppDefinition appDef = AppUtil.getCurrentAppDefinition();

                        Thread excelDownloadThread = new PluginThread(new Runnable() {
                            public void run() {

                                File excelFile = new File(excelFolder, excelFileName);
                                try {
                                    AppUtil.setCurrentAppDefinition(appDef);
                                    dataList.setUseSession(false);
                                    /*
                                     * Previous implementation (preserved in
                                     * DownloadCsvOrExcelUtil#getExcel) loaded
                                     * up to 50,000,000 rows and the full
                                     * XSSFWorkbook in heap. The new method
                                     * fetches all-row exports in batches and
                                     * uses SXSSFWorkbook temporary files.
                                     */
                                    DataListCollection selectedRows = getSelectedRowsForExport(dataList, rowKeys);
                                    DownloadCsvOrExcelUtil.generateStreamingExcelFile(dataList, selectedRows, rowKeys, excelFile, false, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt, exportNumeric, selectedNumericColumn, headerStyle, batchSize, progress::update);

                                    if (storeToForm) {
                                        progress.stage("storing");
                                        DownloadCsvOrExcelUtil.storeGeneratedFileToFormChecked(excelFile, formDefId, fileFieldId);
                                    }
                                    new File(excelFile.getPath() + ".completed").createNewFile();
                                    progress.stage("ready");

                                } catch (Exception e) {
                                    LogUtil.error(getClassName(), e, "Failed in file creation process");
                                    try { progress.stage("failed"); } catch (IOException statusError) {
                                        LogUtil.error(getClassName(), statusError, "Unable to record export failure");
                                    }
                                }
                            }
                        });
                        excelDownloadThread.setDaemon(true);
                        excelDownloadThread.start();

                        AppDefinition appDefination = AppUtil.getCurrentAppDefinition();
                        String url = "/jw/web/json/app/" + appDefination.getAppId() + "/" + appDefination.getVersion()
                                + "/plugin/org.joget.marketplace.DownloadCsvOrExcelDatalistAction/service?uniqueId=" + uniqueId
                                + "&filename=" + java.net.URLEncoder.encode(excelFileName, "UTF-8")
                                + "&storeToForm=" + getPropertyString("storeToForm")
                                + "&downloadBackgroud=" + getPropertyString("downloadBackgroud");
                        result.setUrl(url);

                    } else {
                        // Foreground and store-to-form modes use the same
                        // temporary XLSX file, avoiding an in-memory byte copy.
                        DataListCollection selectedRows = getSelectedRowsForExport(dataList, rowKeys);
                        String excelFileName = renameFile.equalsIgnoreCase("true") ? fileName + ".xlsx" : "report.xlsx";
                        File tempFolder = new File(FileManager.getBaseDirectory(), UuidGenerator.getInstance().getUuid());
                        File excelFile = new File(tempFolder, excelFileName);
                        try {
                            DownloadCsvOrExcelUtil.generateStreamingExcelFile(dataList, selectedRows, rowKeys, excelFile, false, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt, exportNumeric, selectedNumericColumn, headerStyle, batchSize);
                            if (storeToForm) {
                                DownloadCsvOrExcelUtil.storeGeneratedFileToForm(excelFile, formDefId, fileFieldId);
                            } else {
                                DownloadCsvOrExcelUtil.streamExcelFileToResponse(response, excelFile, excelFileName);
                            }
                        } finally {
                            deleteTemporaryExport(excelFile, tempFolder);
                        }
                    }
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
        return result;
    }

    private DataListCollection getDataListRows(DataList dataList, String[] rowKeys, boolean background) {
        DataListCollection dataListRows = null;
        if (rowKeys != null && rowKeys.length > 0) {
            addDataListFilter(dataList, rowKeys);
            dataListRows = dataList.getRows();
        } else {
            if (background) {
                // Legacy all-row behavior retained for reference only. Active
                // streaming call sites never invoke this branch for all rows.
                dataListRows = dataList.getRows(50000000, null);
            } else {
                dataListRows = dataList.getRows(0, 0);
            }
        }
        return dataListRows;
    }

    /**
     * Selected rows retain the legacy filtering behavior. For an all-row
     * export this deliberately returns null; the streaming writer will fetch
     * the datalist in configured-size batches instead of one large collection.
     */
    private DataListCollection getSelectedRowsForExport(DataList dataList, String[] rowKeys) {
        if (rowKeys == null || rowKeys.length == 0) {
            return null;
        }
        return getDataListRows(dataList, rowKeys, false);
    }

    private void deleteTemporaryExport(File excelFile, File tempFolder) {
        if (excelFile.exists() && !excelFile.delete()) {
            LogUtil.warn(getClassName(), "Unable to delete temporary export file: " + excelFile);
        }
        if (tempFolder.isDirectory() && !tempFolder.delete()) {
            LogUtil.warn(getClassName(), "Unable to delete temporary export folder: " + tempFolder);
        }
    }

    public void addDataListFilter(DataList dataList, String[] rowKeys) {
        if (!dataList.isUseSession()) {
            DataListFilterQueryObject filterKeys = new DataListFilterQueryObject();
            filterKeys.setOperator("AND");
            String column = dataList.getBinder().getColumnName(dataList.getBinder().getPrimaryKeyColumnName());
            String query = "";
            for (int i = 0; i < rowKeys.length; i++) {
                if (!query.isEmpty()) {
                    query += ",";
                }
                query += "?";
            }
            filterKeys.setQuery(column + " IN (" + query + ")");
            filterKeys.setValues(rowKeys);
            dataList.addFilterQueryObject(filterKeys);
        }
    }

    @Override
    public void webService(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        boolean roleAnonymous = WorkflowUtil.isCurrentUserInRole(WorkflowUserManager.ROLE_ANONYMOUS);
        if (roleAnonymous) {
            response.sendError(HttpServletResponse.SC_UNAUTHORIZED);
            return;
        }

        response.setHeader("Cache-Control", "no-store");
        String uniqueId = request.getParameter("uniqueId");
        String filename = request.getParameter("filename"); // Servlet parameters are already URL-decoded.
        if (uniqueId == null || !uniqueId.matches("[A-Za-z0-9-]+")
                || filename == null || filename.isEmpty() || filename.contains("/") || filename.contains("\\")) {
            response.sendError(HttpServletResponse.SC_BAD_REQUEST);
            return;
        }
        File folder = new File(FileManager.getBaseDirectory(), uniqueId);
        File file = new File(folder, filename);
        if (!file.getCanonicalFile().getParentFile().equals(folder.getCanonicalFile()) || !folder.isDirectory()) {
            response.sendError(HttpServletResponse.SC_NOT_FOUND);
            return;
        }
        String status = request.getParameter("status");
        if ("progress".equals(status)) {
            response.setContentType("application/json");
            response.setCharacterEncoding("UTF-8");
            response.getWriter().write(BackgroundExportStatus.readJson(folder, file));
        } else if ("generated".equals(status) || "stored".equals(status)) {
            if (!DownloadCsvOrExcelUtil.checkCompletionFlag(file.getPath()) || !file.isFile()) {
                response.sendError(HttpServletResponse.SC_CONFLICT);
                return;
            }
            DownloadCsvOrExcelUtil.streamExcelFileToResponse(response, file, filename);
            deleteCompletedBackgroundExport(file);
        } else {
            response.setContentType("text/html");
            response.setCharacterEncoding("UTF-8");
            try (java.io.InputStream in = getClass().getResourceAsStream("/templates/background-export.html")) {
                if (in == null) { throw new IOException("Missing background export template"); }
                java.io.ByteArrayOutputStream page = new java.io.ByteArrayOutputStream();
                byte[] buffer = new byte[4096];
                int count;
                while ((count = in.read(buffer)) != -1) { page.write(buffer, 0, count); }
                response.getWriter().write(page.toString("UTF-8"));
            }
        }
    }

    /** Remove the completed background artifact after a successful download. */
    private void deleteCompletedBackgroundExport(File excelFile) throws IOException {
        File completionFlag = new File(excelFile.getPath() + ".completed");
        File exportFolder = excelFile.getParentFile();
        java.nio.file.Files.deleteIfExists(new File(exportFolder, "progress.properties").toPath());
        if (excelFile.exists() && !excelFile.delete()) {
            LogUtil.warn(getClassName(), "Unable to delete completed export file: " + excelFile);
        }
        if (completionFlag.exists() && !completionFlag.delete()) {
            LogUtil.warn(getClassName(), "Unable to delete completion flag: " + completionFlag);
        }
        if (exportFolder != null && exportFolder.isDirectory() && !exportFolder.delete()) {
            LogUtil.warn(getClassName(), "Unable to delete completed export folder: " + exportFolder);
        }
    }
}
