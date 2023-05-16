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
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Collection;
import java.util.HashSet;
import java.util.Set;

import org.apache.commons.lang3.math.NumberUtils;


public class DownloadCsvOrExcelDatalistAction extends DataListActionDefault {

    private final DuplicateAndSkip duplicates = new DuplicateAndSkip();

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
        return downloadAs.equalsIgnoreCase("csv");
    }

    public boolean getFooter(){
        String footer = getPropertyString("footerHeader");
        return footer.equalsIgnoreCase("true");
    }

    public boolean includeCustomHeader(){
        String header = getPropertyString("includeCustomHeader");
        return header.equalsIgnoreCase("true");
    }

    public boolean includeCustomFooter(){
        String footer = getPropertyString("includeCustomFooter");
        return footer.equalsIgnoreCase("true");
    }

    @Override
    public DataListActionResult executeAction(DataList dataList, String[] rowKeys) {
        // only allow POST
        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }

        // check for submited rows
        if ( (rowKeys != null && rowKeys.length > 0) || getProperty("downloadAllWhenNoneSelected").equals("true") ) {
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
        String filename = getPropertyString("renameFile").equalsIgnoreCase("true") ? getPropertyString("filename") + ".csv" :"report.csv";
        writeResponse(request, response, bytes, filename, "text/csv");
    }

    protected void downloadExcel(HttpServletRequest request, HttpServletResponse response, DataList dataList, String[] rowKeys) throws ServletException, IOException {
        Workbook workbook = getExcel(dataList, rowKeys);
        String filename = getPropertyString("renameFile").equalsIgnoreCase("true") ? getPropertyString("filename") + ".xlsx" :"report.xlsx";
        writeResponseExcel(request, response, workbook, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\n");
    }

    protected byte[] getCSV(DataList dataList, String[] rowKeys) throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        DataListCollection rows = dataList.getRows();
        HashMap<String, StringBuilder> sb = getLabelAndKey(dataList);
        StringBuilder keySB = sb.get("key");
        StringBuilder headerSB = sb.get("header");

        String[] res = keySB.toString().split(",", 0);
        duplicates.setMap(findDuplicate(res));

        if(includeCustomHeader()){
            outputStream.write((getPropertyString("headerDecorator")+"\n").getBytes());
        }
        outputStream.write((headerSB +"\n").getBytes());

            
        if(rowKeys != null && rowKeys.length > 0){

            //goes through all the datalist row
            for (int x=0; x<rows.size(); x++) {
                //compare with all the rowkeys that have been selected
                for (String rowKey : rowKeys) {
                    //check instance of HashMap if not it will be Formrow
                    boolean boolInstance = rows.get(x) instanceof HashMap;
                    boolean foundRowKey = foundRowKey(boolInstance, rows, x, rowKey);

                    //if no row is found skip
                    if (!foundRowKey) {
                        continue;
                    }

                    Object row = getRow(rows, x);

                    //get the keys and save it
                    printCSV(dataList, outputStream, res, row);
                }
            }

        }else if(getProperty("downloadAllWhenNoneSelected").equals("true")){

            //retrieve all rows
            rows = dataList.getRows(0,0);

            //goes through all the datalist row
            for (int x=0; x<rows.size(); x++) {
                Object row = getRow(rows,x);
                //get the keys and save it
                printCSV(dataList, outputStream, res, row);
            }
        }
        
        if(getFooter()) {
            outputStream.write((headerSB +"\n").getBytes());
        }
        if(includeCustomFooter()){
            outputStream.write((getPropertyString("footerDecorator")+"\n").getBytes());
        }

        return outputStream.toByteArray();
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
        duplicates.setMap(findDuplicate(res));

        if(includeCustomHeader()){
            Cell titleCell = headerRow.createCell(0);
            String headerString = getPropertyString("headerDecorator");
            titleCell.setCellValue(headerString);
            int getNewLine = headerString.split("\r\n|\r|\n").length;
            headerRow.setHeightInPoints((getNewLine * sheet.getDefaultRowHeightInPoints()));
            sheet.autoSizeColumn(2);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, header.length-1));
            rowCounter+=1;
        }

        Row headerColumnRow = sheet.createRow(rowCounter);
        counter = 0;
        for(String myStr: header) {
            Cell headerCell = headerColumnRow.createCell(counter);
            headerCell.setCellValue(myStr);
            counter += 1;
        }

        rowCounter +=1;
        counter = 0;

        if(rowKeys != null && rowKeys.length > 0){
            
            //goes through all the datalist row
            for (int x=0; x<rows.size(); x++) {
                //compare with all the rowkeys that have been selected
                for (int y=0; y<rowKeys.length; y++) {
                    boolean boolInstance = rows.get(x) instanceof HashMap;
                    boolean foundRowKey = foundRowKey(boolInstance,rows,x,rowKeys[y]);

                    if(!foundRowKey) {
                        continue;
                    }

                    printExcel(sheet,rowCounter,counter,rows,x,res,dataList);
                    counter += 1;
                    rowCounter+=1;

                }
            }
            
        }else if(getProperty("downloadAllWhenNoneSelected").equals("true")){
            
            //retrieve all rows
            rows = dataList.getRows(0,0);

            //goes through all the datalist row
            for (int x=0; x<rows.size(); x++) {
                printExcel(sheet,rowCounter,counter,rows,x,res,dataList);
                counter += 1;
                rowCounter+=1;
            }
        }
        
        if(getFooter()) {
            int z = 0;
            Row dataRow = sheet.createRow(rowCounter);
            for (String myStr : header) {
                Cell footerCell = dataRow.createCell(z);
                footerCell.setCellValue(myStr);
                z += 1;
            }
            rowCounter+=1;
        }

        if(includeCustomFooter()){
            Row footerColumnRow = sheet.createRow(rowCounter);
            Cell titleCell = footerColumnRow.createCell(0);
            titleCell.setCellValue(getPropertyString("footerDecorator"));
            sheet.addMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, header.length-1));
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
        int skip = duplicates.getSkipCount(name);
        for (DataListColumn c : columns) {
            if(c.getName().equalsIgnoreCase(name)){

                if(duplicates.checkKey(name)) {
                    if(duplicates.skipCountLessThenDuplicate(name)) {
                        duplicates.addSkipCount(name);
                    }
                    if(skip != 0) {
                        skip -= 1;
                        continue;
                    }
                }

                String value;
                try{
                    value = DataListService.evaluateColumnValueFromRow(o, name).toString();
                    Collection<DataListColumnFormat> formats = c.getFormats();
                    if (formats != null) {
                        for (DataListColumnFormat f : formats) {
                            if (f != null) {
                                value = f.format(dataList, c, o, value);
                                String stripHTML = value.replaceAll("<[^>]*>", "");
                                return stripHTML;
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


    private void printCSV(DataList dataList, ByteArrayOutputStream outputStream, String[] res, Object row) throws IOException {
        for(String myStr: res) {
            String value = getBinderFormattedValue(dataList,row,myStr);

            outputStream.write(value.trim().getBytes());
            outputStream.write(",".getBytes());
        }
        String outputString = new String(outputStream.toByteArray());
        outputString = outputString.substring(0, outputString.length() - 1);
        outputStream.reset();
        outputStream.write(outputString.getBytes());
        outputStream.write("\n".getBytes());
    }

    private void printExcel(Sheet sheet, int rowCounter, int counter, DataListCollection rows, int x, String[] res, DataList dataList) {
        Row dataRow = sheet.createRow(rowCounter);
        Object row = getRow(rows,x);
        int z = 0;
        for(String myStr: res) {
            String value = getBinderFormattedValue(dataList,row,myStr);
            Cell dataRowCell = dataRow.createCell(z);
            if( NumberUtils.isParsable(value) ){
                dataRowCell.setCellValue(Double.parseDouble(value));
            }else{
                dataRowCell.setCellValue(value);
            }
            z += 1;
        }
    }

    private Object getRow(DataListCollection rows, int x) {
        return rows.get(x);
    }

    private boolean foundRowKey(boolean boolInstance, DataListCollection rows, int x, String rowKey) {
        if(boolInstance) {
            return ((HashMap) rows.get(x)).get("id").equals(rowKey);
        } else {
            return ((FormRow) rows.get(x)).get("id").equals(rowKey);
        }
    }
}