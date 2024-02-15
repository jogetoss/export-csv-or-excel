package org.joget.marketplace;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import org.joget.marketplace.DownloadCsvOrExcelDatalistAction;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.*;
import java.util.HashMap;

import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.datalist.model.DataListActionResult;
import org.joget.apps.datalist.service.DataListService;
import org.joget.apps.form.model.FormRow;
import org.joget.commons.util.LogUtil;
import java.net.URL;
import java.util.Map;

import org.joget.apps.app.dao.DatalistDefinitionDao;
import org.joget.apps.app.model.DatalistDefinition;
import org.joget.apps.app.service.AppService;
import org.joget.apps.form.model.FormRowSet;
import org.joget.plugin.base.DefaultApplicationPlugin;
import org.springframework.context.ApplicationContext;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.joget.apps.form.service.FileUtil;
import org.joget.plugin.base.Plugin;
import org.joget.plugin.base.PluginManager;

public class GenerateExcelTool  extends DefaultApplicationPlugin{   
    
    private final static String MESSAGE_PATH = "messages/GenerateExcelTool";

    @Override
    public Object execute(Map props) {
                
        String recordId = getPropertyString("formRecordId");//@recordId@";
        String datalistId = getPropertyString("listId");
        String formDefId = getPropertyString("formDefId");
        String formf = getPropertyString("formField");
        
        LogUtil.info("App - recordID", "Generating Excel File for [" + recordId + "]");
        
        //Joget Configuration
        ApplicationContext ac = AppUtil.getApplicationContext();
        AppService appService = (AppService) ac.getBean("appService");
        
        DataListService dataListService = (DataListService) ac.getBean("dataListService");
        DatalistDefinitionDao datalistDefinitionDao = (DatalistDefinitionDao) ac.getBean("datalistDefinitionDao");
        
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        DatalistDefinition datalistDefinition = datalistDefinitionDao.loadById(datalistId, appDef);

        DataList datalist = dataListService.fromJson(datalistDefinition.getJson());
        
        System.out.println("Datalist: [" + datalist + "]");
        System.out.println("Datalist Action: [" + datalist.getActions() + "]");

        for (DataListAction action : datalist.getActions()) {
                // invoke action
                try{
                    //generate pdf
                    DownloadCsvOrExcelDatalistAction downloadExcelAction = (DownloadCsvOrExcelDatalistAction) action;
                    
                    //
                    Workbook wb = downloadExcelAction.getExcel(datalist, datalist.getRows(), null, false);
                    byte[] byteFile = workbookToByteArray(wb);
                    String path = FileUtil.getUploadPath( appService.getFormTableName(appDef, formDefId) , recordId);
                    String fileName = getPropertyString("fileName") + ".xlsx";
                    final File file = new File(path + fileName);
                    FileUtils.writeByteArrayToFile(file, byteFile);
                    
                    //get original agreement file name
                    FormRowSet set = appService.loadFormData(appDef.getAppId(), appDef.getVersion().toString(), formDefId, recordId);
                    FormRow row = set.get(0);   
                    
                    if(!row.get(getPropertyString("formField")).toString().isEmpty()){
                        row.put(getPropertyString("formField"), row.get(getPropertyString("formField")).toString() + ";" + fileName);
                    }else{
                        row.put(getPropertyString("formField"), fileName);
                    }
                    set.remove(0);
                    set.add(0, row);
                    appService.storeFormData(appDef.getAppId(), appDef.getVersion().toString(), formDefId, set, recordId);
                    System.out.println("Excel File Generated Successfully for [" + recordId + "]");

                } catch (Exception ex) {
                    LogUtil.error("App - recordID", ex, "Failed to generate Excel File for [" + recordId + "]");
                }
                break;
    }
        return null;
    }
    
    private static byte[] workbookToByteArray(Workbook workbook){
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            workbook.write(bos);
            return bos.toByteArray();
        }catch(Exception e){
            return null;
        }
    }

    public String getName() {
        return "Generate Excel Tool";
    }

    public String getVersion() {
        return "8.0.0";
    }

    public String getDescription() {
        return AppPluginUtil.getMessage("org.joget.GenerateExcelTool.pluginDesc" , getClassName() , MESSAGE_PATH);
    }

    public String getLabel() {
        return AppPluginUtil.getMessage("org.joget.GenerateExcelTool.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    public String getClassName() {
        return getClass().getName();
    }

    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/generateExcelTool.json", null, true, MESSAGE_PATH);
    }

}
