package org.joget.marketplace;

import org.apache.poi.ss.usermodel.Workbook;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListCollection;
import org.joget.apps.datalist.service.DataListService;
import org.joget.apps.form.service.FileUtil;
import org.joget.commons.util.LogUtil;
import org.joget.workflow.util.WorkflowUtil;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;

import javax.servlet.http.HttpServletRequest;
import java.nio.file.Files;
import java.util.HashMap;
import java.util.Map;
import java.io.File;
import java.io.IOException;

import org.joget.apps.app.dao.DatalistDefinitionDao;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.model.DatalistDefinition;
import org.joget.marketplace.util.DownloadCsvOrExcelUtil;
import org.joget.plugin.base.DefaultApplicationPlugin;

public class DownloadCsvOrExcelTool extends DefaultApplicationPlugin {

    private final static String MESSAGE_PATH = "messages/DownloadCSVOrExcelTool";

    @Override
    public String getName() {
        return AppPluginUtil.getMessage("org.joget.DownloadCSVOrExcelTool.pluginLabel", getClassName(), MESSAGE_PATH);
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
        return AppPluginUtil.getMessage("org.joget.DownloadCSVOrExcelTool.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getDescription() {
        //support i18n
        return AppPluginUtil.getMessage("org.joget.DownloadCSVOrExcelTool.pluginDesc", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/DownloadCSVOrExcelTool.json", null, true, MESSAGE_PATH);
    }

    public boolean getDownloadAs() {
        String downloadAs = getPropertyString("downloadAs");
        return downloadAs.equalsIgnoreCase("csv");
    }

     @Override
    public Object execute(Map properties) {
        String renameFile = getPropertyString("renameFile");
        String fileName = getPropertyString("filename");
        String delimiter = getPropertyString("delimiter");
        String headerDecorator = getPropertyString("headerDecorator"); 
        String downloadAllWhenNoneSelected = "true"; 
        String footerDecorator = getPropertyString("footerDecorator");
        String includeCustomHeader = getPropertyString("includeCustomHeader"); 
        String footerHeader = getPropertyString("footerHeader"); 
        String includeCustomFooter = getPropertyString("includeCustomFooter");
        String formDefId = getPropertyString("formDefId");
        String fileFieldId = getPropertyString("fileFieldId");
        String pathOptions = getPropertyString("pathOptions");
        String exportImages = getPropertyString("exportImages");;
        String filePath = getPropertyString("filePath");

        ApplicationContext ac = AppUtil.getApplicationContext();
        AppService appService = (AppService) ac.getBean("appService");
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();

        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }
 
        DataList dataList = getDataList(getPropertyString("listDefId"));
        DataListCollection rows = dataList.getRows();
        String[] rowKeys = null;
        String recordId = getPropertyString("recordId");
        if (recordId == null || recordId.equals("")) {
            // get all rowkeys
            rowKeys = new String[rows.size()];
            for (int i = 0; i < rows.size(); i++) {
                Object idObj = ((HashMap) rows.get(i)).get("id");
                rowKeys[i] = idObj != null ? idObj.toString() : null;
            }
        } else {
            // specified row
            rowKeys = new String[] {recordId};
        }

        if ("FILE_PATH".equalsIgnoreCase(pathOptions)) {
            File outputFile = null;
            try {
                if(getDownloadAs()){
                    String filename = renameFile.equalsIgnoreCase("true") ? fileName + ".csv" : "report.csv";
                    outputFile = DownloadCsvOrExcelUtil.generateCSVFile(dataList, rows, rowKeys, renameFile, filePath + "/" + filename, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter);
                } else {
                    Workbook workbook = DownloadCsvOrExcelUtil.getExcel(dataList, rows, rowKeys, false, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages);
                    String filename =renameFile.equalsIgnoreCase("true") ? fileName + ".xlsx" : "report.xlsx";
                    outputFile = DownloadCsvOrExcelUtil.generateExcelOutputFile(workbook, filePath + "/" + filename);
                }
                if (outputFile.exists()) {
                    LogUtil.info(getClassName(), "File saved to: " + filePath);
                } 
            } catch (Exception e){
                 LogUtil.error(getClassName(), e, e.getMessage());   
            }
        } else if ("FORM_FIELD".equalsIgnoreCase(pathOptions)) {
            if(getDownloadAs()){
                DownloadCsvOrExcelUtil.storeCSVToForm(request, dataList, rows, rowKeys, renameFile, fileName, formDefId, fileFieldId, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader,  footerHeader, includeCustomFooter);
            } else {
                Workbook workbook = DownloadCsvOrExcelUtil.getExcel(dataList, rows, rowKeys, false, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages);
                DownloadCsvOrExcelUtil.storeExcelToForm(workbook, getPropertyString("filename") + ".xlsx", renameFile, formDefId, fileFieldId);
            }
            LogUtil.info(getClassName(), "File saved to form");
        }
        
        
        return null;
    }

    protected static DataList getDataList(String datalistId) throws BeansException {
        ApplicationContext ac = AppUtil.getApplicationContext();
        DataListService dataListService = (DataListService) ac.getBean("dataListService");
        DatalistDefinitionDao datalistDefinitionDao = (DatalistDefinitionDao) ac.getBean("datalistDefinitionDao");
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        DatalistDefinition datalistDefinition = datalistDefinitionDao.loadById(datalistId, appDef);
        DataList datalist = null;
        
        if (datalistDefinition != null) {
            datalist = dataListService.fromJson(datalistDefinition.getJson());
        }
        
        return datalist;
    }
}