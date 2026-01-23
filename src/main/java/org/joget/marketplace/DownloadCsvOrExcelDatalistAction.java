package org.joget.marketplace;

import org.apache.commons.lang.ArrayUtils;
import org.apache.poi.ss.usermodel.Workbook;
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
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.joget.apps.app.model.AppDefinition;
import org.joget.commons.util.PluginThread;
import org.joget.commons.util.UuidGenerator;
import org.joget.marketplace.util.DownloadCsvOrExcelUtil;
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
                    DataListCollection dataListRows = getDataListRows(dataList, rowKeys, false);

                    // Check if storeToForm is enabled; skip download if true
                    if (!storeToForm) {
                        DownloadCsvOrExcelUtil.downloadCSV(request, response, dataList, dataListRows, rowKeys, renameFile, fileName, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportEncrypt);
                    }

                    // Store CSV to form if enabled (separate from download)
                    if (storeToForm) {
                        DownloadCsvOrExcelUtil.storeCSVToForm(request, dataList, dataListRows, rowKeys, renameFile, fileName, formDefId, fileFieldId, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportEncrypt);
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
                        AppDefinition appDef = AppUtil.getCurrentAppDefinition();

                        Thread excelDownloadThread = new PluginThread(new Runnable() {
                            public void run() {
                                AppUtil.setCurrentAppDefinition(appDef);
                                dataList.setUseSession(false);
                                DataListCollection rows = getDataListRows(dataList, rowKeys, true);
                                //DataListCollection rows = dataList.getRows(50000000, null);
                                Workbook workbook = DownloadCsvOrExcelUtil.getExcel(dataList, rows, rowKeys, true, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt, exportNumeric, selectedNumericColumn);
                                String filePath = excelFolder.getPath() + File.separator + excelFileName;

                                try {
                                    try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                                        workbook.write(fileOut);
                                    } catch (IOException e) {
                                        LogUtil.error(getClassName(), e, e.getMessage());
                                    }

                                    if (storeToForm) {
                                        new File(filePath + ".completed").createNewFile();
                                        DownloadCsvOrExcelUtil.storeExcelToForm(workbook, excelFileName, renameFile, formDefId, fileFieldId);
                                    }

                                    if (!storeToForm) {
                                        new File(filePath + ".completed").createNewFile();
                                    }

                                } catch (Exception e) {
                                    LogUtil.error(getClassName(), e, "Failed in file creation process");
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
                        // not in the backgroud, get the rows
                        DataListCollection rows = getDataListRows(dataList, rowKeys, false);

                        if (storeToForm) {
                            Workbook workbook = DownloadCsvOrExcelUtil.getExcel(dataList, rows, rowKeys, false, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt, exportNumeric, selectedNumericColumn);
                            DownloadCsvOrExcelUtil.storeExcelToForm(workbook, getPropertyString("filename") + ".xlsx", renameFile, formDefId, fileFieldId);

                        }

                        if (!storeToForm) {
                            DownloadCsvOrExcelUtil.downloadExcel(request, response, dataList, rows, rowKeys, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, renameFile,  fileName, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt, exportNumeric, selectedNumericColumn);
                        }
                    }
                }
            } catch (ServletException e) {
                LogUtil.error(getClassName(), e, "Fail to generate Excel or CSV for " + ArrayUtils.toString(rowKeys));
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
                dataListRows = dataList.getRows(50000000, null);
            } else {
                dataListRows = dataList.getRows(0, 0);
            }
        }
        return dataListRows;
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

        String uniqueId = request.getParameter("uniqueId");
        String filename = request.getParameter("filename");
        try {
            filename = java.net.URLDecoder.decode(filename, "UTF-8");
        } catch (Exception e) {
            LogUtil.error(getClassName(), e, "Failed to decode filename");
        }
        String status = request.getParameter("status");
        boolean storeToForm = Boolean.parseBoolean(request.getParameter("storeToForm"));  // Assuming this parameter exists
        boolean downloadBackground = Boolean.parseBoolean(request.getParameter("downloadBackgroud"));  // Assuming this parameter exists

        if (uniqueId != null && !uniqueId.isEmpty()) {
            if (filename != null && !filename.isEmpty()) {
                // check for the flag file
                String path = FileManager.getBaseDirectory() + File.separator + uniqueId + File.separator + filename;
                boolean fileGenerated = DownloadCsvOrExcelUtil.checkCompletionFlag(path);

                // Check if both storeToForm and downloadBackground are true
                if (storeToForm && downloadBackground) {

                    if (fileGenerated && "stored".equalsIgnoreCase(status)) {
                        File file = new File(path);
                        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                        response.setHeader("Content-Disposition", "attachment; filename=" + file.getName() + "");
                        OutputStream outputStream;
                        try (InputStream inputStream = new FileInputStream(file)) {
                            outputStream = response.getOutputStream();
                            byte[] buffer = new byte[4096];
                            int bytesRead;
                            while ((bytesRead = inputStream.read(buffer)) != -1) {
                                outputStream.write(buffer, 0, bytesRead);
                                outputStream.flush();
                            }
                        }
                        outputStream.close();

                    } else if (fileGenerated) {

                        //datalist.downloadCSVOrExcel.downloadCompleteMessage
                        String message = AppPluginUtil.getMessage("datalist.downloadCSVOrExcel.downloadCompleteMessageStoretoform", getClassName(), MESSAGE_PATH);
                        String flagParam = "status=generated"; // Replace with your desired parameter and value
                        String currentURL = request.getRequestURL().toString();

                        if (request.getQueryString() != null) {
                            currentURL += "?" + request.getQueryString() + "&" + flagParam;
                        } else {
                            currentURL += "?" + flagParam;
                        }

                        String javascript = "<div style='text-align:center;'><h2>" + message + "</h2></div>"
                                + "<script>\n"
                                + "  setTimeout(function() {\n"
                                + "    location.href = '" + currentURL + "';\n"
                                + "  }, 3000); // Redirect after 3 seconds\n"
                                + "</script>";

                        response.setContentType("text/html");
                        response.setCharacterEncoding("UTF-8");
                        try {
                            response.getWriter().write(javascript);
                            response.flushBuffer();
                        } catch (IOException ex) {
                            LogUtil.error(getClassName(), ex, ex.getMessage());
                        }

                    } else {
                        String message = AppPluginUtil.getMessage("datalist.downloadCSVOrExcel.backgroundMessage", getClassName(), MESSAGE_PATH);
                        String javascript = "<marquee width=\"60%\" direction=\"left\" height=\"100px\"><h2>" + message + "</h2></marquee>"
                                + "<script>\n"
                                + "  setTimeout(function() {\n"
                                + "    location.reload();\n"
                                + "  }, 10000); // Reload after 10 seconds\n"
                                + "</script>";

                        response.setContentType("text/html");
                        response.setCharacterEncoding("UTF-8");
                        try {
                            response.getWriter().write(javascript);
                            response.flushBuffer();
                        } catch (IOException ex) {
                            LogUtil.error(getClassName(), ex, ex.getMessage());
                        }

                    }

                } else if (fileGenerated && "generated".equalsIgnoreCase(status)) {
                    File file = new File(path);
                    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    response.setHeader("Content-Disposition", "attachment; filename=" + file.getName() + "");
                    OutputStream outputStream;
                    try (InputStream inputStream = new FileInputStream(file)) {
                        outputStream = response.getOutputStream();
                        byte[] buffer = new byte[4096];
                        int bytesRead;
                        while ((bytesRead = inputStream.read(buffer)) != -1) {
                            outputStream.write(buffer, 0, bytesRead);
                            outputStream.flush();
                        }
                    }
                    outputStream.close();

                } else if (fileGenerated) {

                    //datalist.downloadCSVOrExcel.downloadCompleteMessage
                    String message = AppPluginUtil.getMessage("datalist.downloadCSVOrExcel.downloadCompleteMessage", getClassName(), MESSAGE_PATH);
                    String flagParam = "status=generated"; // Replace with your desired parameter and value
                    String currentURL = request.getRequestURL().toString();

                    if (request.getQueryString() != null) {
                        currentURL += "?" + request.getQueryString() + "&" + flagParam;
                    } else {
                        currentURL += "?" + flagParam;
                    }

                    String javascript = "<div style='text-align:center;'><h2>" + message + "</h2></div>"
                            + "<script>\n"
                            + "  setTimeout(function() {\n"
                            + "    location.href = '" + currentURL + "';\n"
                            + "  }, 3000); // Redirect after 3 seconds\n"
                            + "</script>";

                    response.setContentType("text/html");
                    response.setCharacterEncoding("UTF-8");
                    try {
                        response.getWriter().write(javascript);
                        response.flushBuffer();
                    } catch (IOException ex) {
                        LogUtil.error(getClassName(), ex, ex.getMessage());
                    }

                } else {
                    String message = AppPluginUtil.getMessage("datalist.downloadCSVOrExcel.backgroundMessage", getClassName(), MESSAGE_PATH);
                    String javascript = "<marquee width=\"60%\" direction=\"left\" height=\"100px\"><h2>" + message + "</h2></marquee>"
                            + "<script>\n"
                            + "  setTimeout(function() {\n"
                            + "    location.reload();\n"
                            + "  }, 10000); // Reload after 10 seconds\n"
                            + "</script>";

                    response.setContentType("text/html");
                    response.setCharacterEncoding("UTF-8");
                    try {
                        response.getWriter().write(javascript);
                        response.flushBuffer();
                    } catch (IOException ex) {
                        LogUtil.error(getClassName(), ex, ex.getMessage());
                    }

                }
            }
        }
    }
}
