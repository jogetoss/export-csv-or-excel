package org.joget.marketplace;

import org.apache.commons.lang.ArrayUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListColumn;
import org.joget.apps.datalist.model.DataListColumnFormat;
import org.joget.apps.datalist.model.DataListActionDefault;
import org.joget.apps.datalist.model.DataListActionResult;
import org.joget.apps.datalist.model.DataListCollection;
import org.joget.apps.datalist.service.DataListService;
import org.joget.apps.form.model.FormRow;
import org.joget.commons.util.LogUtil;
import org.joget.workflow.util.WorkflowUtil;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Collection;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

import org.apache.commons.lang3.math.NumberUtils;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.datalist.model.DataListFilterQueryObject;
import static org.joget.commons.util.FileManager.getBaseDirectory;
import org.joget.commons.util.PluginThread;
import org.joget.plugin.base.PluginWebSupport;
import org.joget.workflow.model.service.WorkflowUserManager;

public class DownloadCsvOrExcelDatalistAction extends DataListActionDefault implements PluginWebSupport {

    private final DuplicateAndSkip duplicates = new DuplicateAndSkip();

    private static Map<String, Object> data;

    private final static String MESSAGE_PATH = "messages/DownloadCSVOrExcelDatalistAction";

    @Override
    public String getName() {
        return "Download CSV or Excel Datalist Action";
    }

    @Override
    public String getVersion() {
        return "8.0.7";
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

    public boolean getFooter() {
        String footer = getPropertyString("footerHeader");
        return footer.equalsIgnoreCase("true");
    }

    public boolean includeCustomHeader() {
        String header = getPropertyString("includeCustomHeader");
        return header.equalsIgnoreCase("true");
    }

    public boolean includeCustomFooter() {
        String footer = getPropertyString("includeCustomFooter");
        return footer.equalsIgnoreCase("true");
    }

    @Override
    public DataListActionResult executeAction(final DataList dataList, String[] rowKeys) {
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
                if (getDownloadAs()) {
                    DataListCollection dataListRows = getDataListRows(dataList, rowKeys, false);
                    downloadCSV(request, response, dataList, dataListRows, rowKeys);
                } else {
                    String downloadBackgroud = getPropertyString("downloadBackgroud");
                    if ("true".equalsIgnoreCase(downloadBackgroud)) {
                        String uniqueId = UUID.randomUUID().toString();
                        String excelFileName = getPropertyString("renameFile").equalsIgnoreCase("true") ? getPropertyString("filename") + ".xlsx" : "report.xlsx";
                        File excelFolder = new File(getBaseDirectory(), uniqueId);
                        if (!excelFolder.isDirectory()) {
                            //create directories if not exist
                            new File(getBaseDirectory(), uniqueId).mkdirs();
                        }
                        final AppDefinition appDef = AppUtil.getCurrentAppDefinition();
                        Thread excelDownloadThread = new PluginThread(new Runnable() {
                            public void run() {
                                AppUtil.setCurrentAppDefinition(appDef);
                                dataList.setUseSession(false);
                                DataListCollection rows = getDataListRows(dataList, rowKeys, true);
                                //DataListCollection rows = dataList.getRows(50000000, null);
                                Workbook workbook = getExcel(dataList, rows, rowKeys, true);
                                String filePath = excelFolder.getPath() + File.separator + excelFileName;
                                try ( FileOutputStream fileOut = new FileOutputStream(filePath)) {
                                    workbook.write(fileOut);
                                } catch (IOException e) {
                                    LogUtil.error(getClassName(), e, e.getMessage());
                                }
                            }
                        });

                        excelDownloadThread.setDaemon(true);
                        excelDownloadThread.start();

                        AppDefinition appDefination = AppUtil.getCurrentAppDefinition();
                        String url = "/jw/web/json/app/" + appDefination.getAppId() + "/" + appDefination.getVersion() + "/plugin/org.joget.marketplace.DownloadCsvOrExcelDatalistAction/service?uniqueId=" + uniqueId + "&filename=" + excelFileName;
                        result.setUrl(url);

                    } else {
                        // not in the backgroud, get the rows
                        DataListCollection rows = getDataListRows(dataList, rowKeys, false);
                        downloadExcel(request, response, dataList, rows, rowKeys);
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

    protected void downloadCSV(HttpServletRequest request, HttpServletResponse response, DataList dataList, DataListCollection dataListRows, String[] rowKeys) throws ServletException, IOException {

        String filename = getPropertyString("renameFile").equalsIgnoreCase("true") ? getPropertyString("filename") + ".csv" : "report.csv";
        String delimiter = getPropertyString("delimiter");
        if (delimiter.isEmpty()) {
            delimiter = ",";
        }
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=" + filename + "");

        try ( OutputStream outputStream = response.getOutputStream()) {
            PrintWriter writer = new PrintWriter(outputStream);
            streamCSV(request, response, writer, dataList, dataListRows, rowKeys, delimiter);
            writer.flush();  // Flush any remaining buffered data
            outputStream.flush();  // Flush the output stream
            writer.close();
        }
    }

    protected void streamCSV(HttpServletRequest request, HttpServletResponse response, PrintWriter writer, DataList dataList, DataListCollection dataListRows, String[] rowKeys, String delimiter) throws IOException {

        HashMap<String, StringBuilder> labelAndKeys = getLabelAndKey(dataList);
        StringBuilder keySB = labelAndKeys.get("key");
        StringBuilder headerSB = labelAndKeys.get("header");

        String[] keys = keySB.toString().split(",", 0);
        duplicates.setMap(findDuplicate(keys));

        if (includeCustomHeader()) {
            writer.write((getPropertyString("headerDecorator") + "\n"));
        }

        if (delimiter != null && !delimiter.isEmpty()) {
            String replacedString = headerSB.toString().replace(",", delimiter);
            headerSB.setLength(0);
            headerSB.append(replacedString);
        }

        writer.write((headerSB + "\n"));

        if (rowKeys != null && rowKeys.length > 0) {
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
                    writeCSVContents(dataList, null, keys, row, writer, delimiter);
                }
            }

        } else if (getProperty("downloadAllWhenNoneSelected").equals("true")) {
            for (int x = 0; x < dataListRows.size(); x++) {
                Object row = getRow(dataListRows, x);
                //get the keys and save it
                writeCSVContents(dataList, null, keys, row, writer, delimiter);
            }
        }
        if (getFooter()) {
            writer.write("\n");
            writer.write((headerSB + "\n"));
        }
        if (includeCustomFooter()) {
            writer.write("\n");
            writer.write((getPropertyString("footerDecorator") + "\n"));
        }
    }

    private void writeCSVContents(DataList dataList, ByteArrayOutputStream outputStream, String[] keys, Object row, PrintWriter writer, String delimiter) throws IOException {
        StringBuilder stringBuilder = new StringBuilder();
        for (String value : keys) {
            String formattedValue = getBinderFormattedValue(dataList, row, value);

            if (formattedValue != null && formattedValue.contains(delimiter)) {
                formattedValue = "\"" + formattedValue + "\"";
            }

            stringBuilder.append(formattedValue);
            stringBuilder.append(delimiter);
        }

        if (stringBuilder.length() > 0) {
            stringBuilder.setLength(stringBuilder.length() - 1);
        }
        String value = stringBuilder.toString();
        writer.write("\r\n");
        writer.write(value);
        writer.flush();
    }

    protected Workbook getExcel(DataList dataList, DataListCollection rows, String[] rowKeys, boolean background) {
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

        if (includeCustomHeader()) {
            Cell titleCell = headerRow.createCell(0);
            String headerString = getPropertyString("headerDecorator");
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

        if (rowKeys != null && rowKeys.length > 0) {
            for (int x = 0; x < rows.size(); x++) {
                //compare with all the rowkeys that have been selected
                for (int y = 0; y < rowKeys.length; y++) {
                    boolean boolInstance = rows.get(x) instanceof HashMap;
                    boolean foundRowKey = foundRowKey(boolInstance, rows, x, rowKeys[y]);

                    if (!foundRowKey) {
                        continue;
                    }
                    printExcel(sheet, rowCounter, counter, rows, x, res, dataList);
                    counter += 1;
                    rowCounter += 1;
                }
            }

        } else if (getProperty("downloadAllWhenNoneSelected").equals("true")) {
            for (int x = 0; x < rows.size(); x++) {
                printExcel(sheet, rowCounter, counter, rows, x, res, dataList);
                counter += 1;
                rowCounter += 1;
            }
        }

        if (getFooter()) {
            int z = 0;
            Row dataRow = sheet.createRow(rowCounter);
            for (String myStr : header) {
                Cell footerCell = dataRow.createCell(z);
                footerCell.setCellValue(myStr);
                z += 1;
            }
            rowCounter += 1;
        }

        if (includeCustomFooter()) {
            Row footerColumnRow = sheet.createRow(rowCounter);
            Cell titleCell = footerColumnRow.createCell(0);
            titleCell.setCellValue(getPropertyString("footerDecorator"));
            
            if (header.length >= 2) {
                sheet.addMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, header.length - 1));
            }
        }

        return workbook;

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
        String status = request.getParameter("status");
        if (uniqueId != null && !uniqueId.isEmpty()) {
            if (filename != null && !filename.isEmpty()) {
                // check for the flag file
                String path = getBaseDirectory() + File.separator + uniqueId + File.separator + filename;
                boolean fileGenerated = checkCompletionFlag(path);
                if (fileGenerated && "generated".equalsIgnoreCase(status)) {
                    File file = new File(path);
                    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    response.setHeader("Content-Disposition", "attachment; filename=" + file.getName() + "");
                    OutputStream outputStream;
                    try ( InputStream inputStream = new FileInputStream(file)) {
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

    private boolean checkCompletionFlag(String path) {
        File flagFile = new File(path);
        return flagFile.exists();
    }

    protected void downloadExcel(HttpServletRequest request, HttpServletResponse response, DataList dataList, DataListCollection dataListRows, String[] rowKeys) throws ServletException, IOException {
        Workbook workbook = getExcel(dataList, dataListRows, rowKeys, false);
        String filename = getPropertyString("renameFile").equalsIgnoreCase("true") ? getPropertyString("filename") + ".xlsx" : "report.xlsx";
        writeResponseExcel(request, response, workbook, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\n");
    }

    private HashMap<String, Integer> findDuplicate(String[] keySB) {
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

    protected void writeResponseExcel(HttpServletRequest request, HttpServletResponse response, Workbook workbook, String filename, String contentType) throws IOException, ServletException {
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

    protected String getBinderFormattedValue(DataList dataList, Object o, String name) {
        DataListColumn[] columns = dataList.getColumns();
        int skip = duplicates.getSkipCount(name);
        for (DataListColumn c : columns) {
            if (c.getName().equalsIgnoreCase(name)) {

                if (duplicates.checkKey(name)) {
                    if (duplicates.skipCountLessThenDuplicate(name)) {
                        duplicates.addSkipCount(name);
                    }
                    if (skip != 0) {
                        skip -= 1;
                        continue;
                    }
                }

                String value;
                try {
                    value = DataListService.evaluateColumnValueFromRow(o, name).toString();
                    Collection<DataListColumnFormat> formats = c.getFormats();
                    if (formats != null) {
                        for (DataListColumnFormat f : formats) {
                            if (f != null) {
                                value = f.format(dataList, c, o, value);
                                String stripHTML = value.replaceAll("<[^>]*>", "");
                                return stripHTML;
                            } else {
                                return value;
                            }
                        }
                    } else {
                        return value;
                    }
                } catch (Exception ex) {

                }
            }
        }
        return "";
    }

    protected HashMap<String, StringBuilder> getLabelAndKey(DataList dataList) {
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

    private void printExcel(Sheet sheet, int rowCounter, int counter, DataListCollection rows, int x, String[] res, DataList dataList) {
        Row dataRow = sheet.createRow(rowCounter);
        Object row = getRow(rows, x);
        int z = 0;
        for (String myStr : res) {
            String value = getBinderFormattedValue(dataList, row, myStr);
            Cell dataRowCell = dataRow.createCell(z);
            if (NumberUtils.isParsable(value)) {
                dataRowCell.setCellValue(Double.parseDouble(value));
            } else {
                dataRowCell.setCellValue(value);
            }
            z += 1;
        }
    }

    private Object getRow(DataListCollection rows, int x) {
        return rows.get(x);
    }

    private boolean foundRowKey(boolean boolInstance, DataListCollection rows, int x, String rowKey) {
        if (boolInstance) {
            return ((HashMap) rows.get(x)).get("id").equals(rowKey);
        } else {
            return ((FormRow) rows.get(x)).get("id").equals(rowKey);
        }
    }

}
