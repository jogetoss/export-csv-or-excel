package org.joget.marketplace;

import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListCollection;
import org.joget.apps.datalist.model.DataListFilterQueryObject;
import org.joget.apps.datalist.service.DataListService;
import org.joget.commons.util.LogUtil;
import org.joget.commons.util.FileManager;
import org.joget.commons.util.UuidGenerator;
import org.joget.workflow.util.WorkflowUtil;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;

import javax.servlet.http.HttpServletRequest;
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
        final int batchSize = DownloadCsvOrExcelUtil.getDataBatchSize(getPropertyString("dataBatchSize"));
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
        String exportEncrypt = getPropertyString("exportEncrypt");
        String exportNumeric = getPropertyString("exportNumeric");
        Object[] selectedNumericColumn = (Object[]) properties.get("selectedNumericColumn");
        final DownloadCsvOrExcelUtil.HeaderStyle headerStyle = DownloadCsvOrExcelUtil.HeaderStyle.of(
                getPropertyString("headerBackgroundColor"), getPropertyString("headerFontColor"),
                getPropertyString("headerBold"), getPropertyString("headerItalic"),
                getPropertyString("headerFontName"), getPropertyString("headerFontSize"),
                getPropertyString("headerAlignment"));

        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }
 
        DataList dataList = getDataList(getPropertyString("listDefId"));
        String[] rowKeys = null;
        DataListCollection selectedRows = null;
        String recordId = getPropertyString("recordId");
        if (recordId != null && !recordId.isEmpty()) {
            // Preserve the original record-id behavior. Only this small
            // selected-row path is preloaded; all-row exports are now paged.
            rowKeys = new String[] {recordId};
            addRecordIdFilter(dataList, rowKeys);
            selectedRows = dataList.getRows();
        }

        if ("FILE_PATH".equalsIgnoreCase(pathOptions)) {
            File outputFile = null;
            try {
                if(getDownloadAs()){
                    String filename = renameFile.equalsIgnoreCase("true") ? fileName + ".csv" : "report.csv";
                    outputFile = DownloadCsvOrExcelUtil.generateStreamingCSVFile(dataList, selectedRows, rowKeys, new File(filePath, filename), true, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportEncrypt, batchSize);
                } else {
                    String filename =renameFile.equalsIgnoreCase("true") ? fileName + ".xlsx" : "report.xlsx";
                    outputFile = DownloadCsvOrExcelUtil.generateStreamingExcelFile(dataList, selectedRows, rowKeys, new File(filePath, filename), true, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt, exportNumeric, selectedNumericColumn, headerStyle, batchSize);
                }
                if (outputFile.exists()) {
                    LogUtil.info(getClassName(), "File saved to: " + filePath);
                } 
            } catch (Exception e){
                 LogUtil.error(getClassName(), e, e.getMessage());   
            }
        } else if ("FORM_FIELD".equalsIgnoreCase(pathOptions)) {
            String generatedName = renameFile.equalsIgnoreCase("true") ? fileName + (getDownloadAs() ? ".csv" : ".xlsx") : (getDownloadAs() ? "report.csv" : "report.xlsx");
            File tempFolder = new File(FileManager.getBaseDirectory(), UuidGenerator.getInstance().getUuid());
            File generatedFile = new File(tempFolder, generatedName);
            try {
                if (getDownloadAs()) {
                    DownloadCsvOrExcelUtil.generateStreamingCSVFile(dataList, selectedRows, rowKeys, generatedFile, false, delimiter, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportEncrypt, batchSize);
                } else {
                    DownloadCsvOrExcelUtil.generateStreamingExcelFile(dataList, selectedRows, rowKeys, generatedFile, false, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt, exportNumeric, selectedNumericColumn, headerStyle, batchSize);
                }
                DownloadCsvOrExcelUtil.storeGeneratedFileToForm(generatedFile, formDefId, fileFieldId);
            } catch (IOException e) {
                LogUtil.error(getClassName(), e, "Failed to generate file for form storage");
            } finally {
                deleteTemporaryExport(generatedFile, tempFolder);
            }
            LogUtil.info(getClassName(), "File saved to form");
        }
        
        
        return null;
    }

    private void deleteTemporaryExport(File generatedFile, File tempFolder) {
        if (generatedFile.exists() && !generatedFile.delete()) {
            LogUtil.warn(getClassName(), "Unable to delete temporary export file: " + generatedFile);
        }
        if (tempFolder.isDirectory() && !tempFolder.delete()) {
            LogUtil.warn(getClassName(), "Unable to delete temporary export folder: " + tempFolder);
        }
    }

    private void addRecordIdFilter(DataList dataList, String[] rowKeys) {
        if (!dataList.isUseSession()) {
            DataListFilterQueryObject filter = new DataListFilterQueryObject();
            filter.setOperator("AND");
            String column = dataList.getBinder().getColumnName(dataList.getBinder().getPrimaryKeyColumnName());
            filter.setQuery(column + " IN (?)");
            filter.setValues(rowKeys);
            dataList.addFilterQueryObject(filter);
        }
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
