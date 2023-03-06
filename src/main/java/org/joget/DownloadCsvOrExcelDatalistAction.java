package org.joget;

import org.apache.commons.lang.ArrayUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

import org.joget.commons.util.LogUtil;
import org.joget.workflow.util.WorkflowUtil;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Collection;


public class DownloadCsvOrExcelDatalistAction extends DataListActionDefault {

    private final static String MESSAGE_PATH = "messages/DownloadCSVOrExcelDatalistAction";

    public String getName() {
        return "Download CSV or Excel Datalist Action";
    }
    public String getVersion() {
        return "8.0.0";
    }

    public String getClassName() {
        return getClass().getName();
    }

    public String getLabel() {
        //support i18n
        return AppPluginUtil.getMessage("org.joget.DownloadCSVOrExcelDatalistAction.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    public String getDescription() {
        //support i18n
        return AppPluginUtil.getMessage("org.joget.DownloadCSVOrExcelDatalistAction.pluginDesc", getClassName(), MESSAGE_PATH);
    }

    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/DownloadCSVOrExcelDatalistAction.json", null, true, MESSAGE_PATH);
    }

    public String getLinkLabel() {
        return getPropertyString("label"); //get label from configured properties options
    }

    public String getHref() {
        return getPropertyString("href"); //Let system to handle to post to the same page
    }

    public String getTarget() {
        return "post";
    }

    public String getHrefParam() {
        return getPropertyString("hrefParam");  //Let system to set the parameter to the checkbox name
    }

    public String getHrefColumn() {
        String recordIdColumn = getPropertyString("recordIdColumn"); //get column id from configured properties options
        if ("id".equalsIgnoreCase(recordIdColumn) || recordIdColumn.isEmpty()) {
            return getPropertyString("hrefColumn"); //Let system to set the primary key column of the binder
        } else {
            return recordIdColumn;
        }
    }

    public String getConfirmation() {
        return getPropertyString("confirmation"); //get confirmation from configured properties options
    }

    public boolean getDownloadAs() {
        String downloadAs = getPropertyString("downloadAs");
        if(downloadAs.equalsIgnoreCase("csv")) {
            return true;
        } else {
            return false;
        }
    }

    public boolean getHeaderLabel(){
        String downloadAs = getPropertyString("repeatHeaderLabel");
        if(downloadAs.equalsIgnoreCase("true")) {
            return true;
        } else {
            return false;
        }
    }

    @Override
    public DataListActionResult executeAction(DataList dataList, String[] rowKeys) {
        // only allow POST
        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }

        // check for submited rows
        if (rowKeys != null && rowKeys.length > 0) {
            try {
                //get the HTTP Response
                HttpServletResponse response = WorkflowUtil.getHttpServletResponse();
                if(getDownloadAs()) {
                    downloadCSV(request, response, dataList, rowKeys);
                } else {
                    downloadExcel(request, response, dataList, rowKeys);
                }
            } catch (ServletException e) {
                LogUtil.error(getClassName(), e, "Fail to generate Excel or CSV for " + ArrayUtils.toString(rowKeys));
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }

        //return null to do nothing
        return null;
    }

    protected void downloadCSV(HttpServletRequest request, HttpServletResponse response, DataList dataList, String[] rowKeys) throws ServletException, IOException {
        byte[] bytes = getCSV(dataList, rowKeys);
        writeResponse(request, response, bytes, "report.csv", "text/csv");
    }

    protected void downloadExcel(HttpServletRequest request, HttpServletResponse response, DataList dataList, String[] rowKeys) throws ServletException, IOException {
        Workbook workbook = getExcel(dataList, rowKeys);
        writeResponseExcel(request, response, workbook, "report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\n");
    }

    protected byte[] getCSV(DataList dataList, String[] rowKeys) throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        DataListCollection rows = dataList.getRows();
        HashMap<String, StringBuilder> sb = getLabelAndKey(dataList);
        StringBuilder keySB = sb.get("key");
        StringBuilder headerSB = sb.get("header");
        int counter = 0;

        String[] res = keySB.toString().split(",", 0);
        outputStream.write((headerSB +"\n").getBytes());

        //goes through all the datalist row
        for (int x=0; x<rows.size(); x++) {
            //compare with all the rowkeys that have been selected
            for (int y=0; y<rowKeys.length; y++) {
                if(((HashMap) rows.get(x)).get("id").equals(rowKeys[y])) {
                    counter += 1;
                    HashMap row = (HashMap) rows.get(x);
                    //get the keys and save it
                    for(String myStr: res) {
                        String value = getBinderFormattedValue(dataList,row,myStr);
                        outputStream.write(value.getBytes());
                        outputStream.write(",".getBytes());
                    }
                    String outputString = new String(outputStream.toByteArray());
                    outputString = outputString.substring(0, outputString.length() - 1);
                    outputStream.reset();
                    outputStream.write(outputString.getBytes());
                    outputStream.write("\n".getBytes());
                }
            }
            if(counter == rowKeys.length) {
                if(getHeaderLabel()) {
                    outputStream.write((headerSB +"\n").getBytes());
                }
                break;
            }
        }

        byte[] bytes = outputStream.toByteArray();
        return bytes;
    }

    protected Workbook getExcel(DataList dataList, String[] rowKeys) {

        DataListCollection rows = dataList.getRows();
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

        for(String myStr: header) {
            Cell headerCell = headerRow.createCell(counter);
            headerCell.setCellValue(myStr);
            counter += 1;
        }
        counter = 0;
        rowCounter +=1;

        //goes through all the datalist row
        for (int x=0; x<rows.size(); x++) {
            //compare with all the rowkeys that have been selected
            for (int y=0; y<rowKeys.length; y++) {
                Row dataRow = sheet.createRow(rowCounter);
                if(((HashMap) rows.get(x)).get("id").equals(rowKeys[y])) {
                    counter += 1;
                    HashMap row = (HashMap) rows.get(x);
                    int z = 0;
                    for(String myStr: res) {
                        String value = getBinderFormattedValue(dataList,row,myStr);
                        Cell dataRowCell = dataRow.createCell(z);
                        dataRowCell.setCellValue(value);
                        z += 1;
                    }
                    rowCounter+=1;
                }
            }
            if(counter == rowKeys.length) {
                if(getHeaderLabel()) {
                    int z = 0;
                    Row dataRow = sheet.createRow(rowCounter);
                    for (String myStr : header) {
                        Cell footerCell = dataRow.createCell(z);
                        footerCell.setCellValue(myStr);
                        z += 1;
                    }
                }
                break;
            }
        }

        return workbook;

    }

    /**
     * Write to response for download
     * @param request
     * @param response
     * @param bytes
     * @param filename
     * @param contentType
     * @throws IOException
     */
    protected void writeResponse(HttpServletRequest request, HttpServletResponse response, byte[] bytes, String filename, String contentType) throws IOException, ServletException {
        OutputStream out = response.getOutputStream();
        try {
            String name = URLEncoder.encode(filename, "UTF8").replaceAll("\\+", "%20");
            response.setHeader("Content-Disposition", "attachment; filename="+name+"; filename*=UTF-8''" + name);
            response.setContentType(contentType+"; charset=UTF-8");

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

    protected void writeResponseExcel(HttpServletRequest request, HttpServletResponse response, Workbook workbook, String filename, String contentType) throws IOException, ServletException {
        OutputStream out = response.getOutputStream();
        try {
            String name = URLEncoder.encode(filename, "UTF8").replaceAll("\\+", "%20");
            response.setHeader("Content-Disposition", "attachment; filename="+name+"; filename*=UTF-8''" + name);
            response.setContentType(contentType+"; charset=UTF-8");

            ByteArrayOutputStream ms=new ByteArrayOutputStream();
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

    protected String getBinderFormattedValue(DataList dataList, Object o, String name){
        DataListColumn[] columns = dataList.getColumns();

        for (DataListColumn c : columns) {
            if(c.getName().equalsIgnoreCase(name)){
                String value;
                try{
                    value = DataListService.evaluateColumnValueFromRow(o, name).toString();
                    Collection<DataListColumnFormat> formats = c.getFormats();
                    if (formats != null) {
                        for (DataListColumnFormat f : formats) {
                            if (f != null) {
                                value = f.format(dataList, c, o, value);
                                return value;
                            }else{
                                return value;
                            }
                        }
                    }else{
                        return value;
                    }
                }catch(Exception ex){

                }
            }
        }
        return "";
    }

    protected HashMap<String, StringBuilder> getLabelAndKey(DataList dataList) {
        HashMap<String, StringBuilder> sb = new HashMap<>();
        StringBuilder headerSB = new StringBuilder();
        StringBuilder keySB = new StringBuilder();

        int counter = 0;
        for (DataListColumn column : dataList.getColumns()) {
            String header = column.getLabel();
            String key = column.getName();

            headerSB.append(header).append(",");
            keySB.append(key).append(",");
        }
        headerSB.setLength(headerSB.length() - 1);
        keySB.setLength(keySB.length() - 1);

        sb.put("header", headerSB);
        sb.put("key", keySB);
        return sb;
    }
}