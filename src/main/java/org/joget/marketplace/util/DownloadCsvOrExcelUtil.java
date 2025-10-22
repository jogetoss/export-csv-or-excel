package org.joget.marketplace.util;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
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
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.util.Collection;
import java.util.HashSet;
import java.util.Set;

import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;

import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.tika.Tika;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppService;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.service.FileUtil;
import org.joget.apps.form.service.FormUtil;
import org.joget.commons.util.UuidGenerator;
import org.springframework.context.ApplicationContext;

public class DownloadCsvOrExcelUtil {

    private final static DuplicateAndSkip duplicates = new DuplicateAndSkip();

    private static Map<String, Object> data;

    private final static String MESSAGE_PATH = "messages/DownloadCSVOrExcelDatalistAction";

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

    public static Workbook getExcel(DataList dataList, DataListCollection rows, String[] rowKeys, boolean background, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator, String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportImages, String exportEncrypt) {
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
            for (int x = 0; x < rows.size(); x++) {
                //compare with all the rowkeys that have been selected
                for (int y = 0; y < rowKeys.length; y++) {
                    boolean boolInstance = rows.get(x) instanceof HashMap;
                    boolean foundRowKey = foundRowKey(boolInstance, rows, x, rowKeys[y]);

                    if (!foundRowKey) {
                        continue;
                    }
                    printExcel(currentAppDef, sheet, rowCounter, counter, rows, x, res, dataList, exportImages, exportEncrypt);
                    counter += 1;
                    rowCounter += 1;
                }
            }

        } else if (downloadAllWhenNoneSelected.equals("true")) {
            for (int x = 0; x < rows.size(); x++) {
                printExcel(currentAppDef, sheet, rowCounter, counter, rows, x, res, dataList, exportImages, exportEncrypt);
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

        return workbook;

    }

    public static boolean checkCompletionFlag(String path) {
        File flagFile = new File(path + ".completed");
        return flagFile.exists();
    }

    public static void downloadExcel(HttpServletRequest request, HttpServletResponse response, DataList dataList, DataListCollection dataListRows, String[] rowKeys, String headerDecorator, String downloadAllWhenNoneSelected, String footerDecorator, String renameFile, String fileName, String includeCustomHeader, String footerHeader, String includeCustomFooter, String exportImages, String exportEncrypt) throws ServletException, IOException {
        Workbook workbook = getExcel(dataList, dataListRows, rowKeys, false, headerDecorator, downloadAllWhenNoneSelected, footerDecorator, includeCustomHeader, footerHeader, includeCustomFooter, exportImages, exportEncrypt);
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

    protected static String getBinderFormattedValue(DataList dataList, Object o, String name, String exportImages, String exportEncrypt) {
        DataListColumn[] columns = dataList.getColumns();
        int skip = duplicates.getSkipCount(name);
        for (DataListColumn c : columns) {
            if (c.getName().equalsIgnoreCase(name)) {

                if ("true".equals(exportImages)) {
                    Collection<DataListColumnFormat> formatsList = c.getFormats();
                    if (formatsList != null && !formatsList.isEmpty()) {
                        DataListColumnFormat firstFormat = null;
                        firstFormat = formatsList.iterator().next();
                        if (firstFormat != null) {
                            String formatterClassName = firstFormat.getClassName();
                            String filename = DataListService.evaluateColumnValueFromRow(o, name).toString();
                            if ("org.joget.apps.datalist.lib.ImageFormatter".equals(formatterClassName)) {
                                // image upload field
                                String formDefId = (String) firstFormat.getProperty("formDefId");
                                String imageSrc = (String) firstFormat.getProperty("imageSrc"); // imageSrc => form
                                if ("form".equals(imageSrc) && filename != null && !filename.isEmpty()) {
                                    return "IMAGE:" + formDefId + ":" + imageSrc + ":" + filename;
                                }
                            } else if ("org.joget.tutorial.FileLinkDatalistFormatter".equals(formatterClassName)) {
                                // file upload field
                                String formDefId = (String) firstFormat.getProperty("formDefId");
                                if (formDefId != null && !formDefId.isEmpty() && filename != null && !filename.isEmpty()) {
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

                String value;
                try {
                    value = DataListService.evaluateColumnValueFromRow(o, name).toString();

                    if (!"true".equals(exportEncrypt)) {
                        value = SecurityUtil.decrypt(value);
                    }

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

    protected static void printExcel(AppDefinition currentAppDef, Sheet sheet, int rowCounter, int counter, DataListCollection rows, int x, String[] res, DataList dataList, String exportImages, String exportEncrypt) {
        Row dataRow = sheet.createRow(rowCounter);
        Object row = getRow(rows, x);
        int z = 0;
        for (String myStr : res) {
            String value = getBinderFormattedValue(dataList, row, myStr, exportImages, exportEncrypt);

            // Check if value contains multiple images (if they're separated by some delimiter)
            if (value.startsWith("IMAGE:") || value.startsWith("FILE:")) {
                // Process image/file - this might consume multiple columns if there are multiple images
                z = processImageOrFile(currentAppDef, sheet, dataRow, rowCounter, rows, x, value, z);
            } else {
                // Process regular data
                Cell dataRowCell = dataRow.createCell(z);
                if (NumberUtils.isParsable(value)) {
                    double numericValue = Double.parseDouble(value);
                    if (isWholeNumber(numericValue)) {
                        // If the numeric value is a whole number, format it to display without decimal points
                        DecimalFormat decimalFormat = new DecimalFormat("#");
                        value = decimalFormat.format(numericValue);
                    } else {
                        // If the numeric value has decimal points, set it directly
                        value = Double.toString(numericValue);
                    }

                    dataRowCell.setCellValue(value);
                } else {
                    dataRowCell.setCellValue(value);
                }
                z += 1;
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
        if (boolInstance) {
            return ((HashMap) rows.get(x)).get("id").equals(rowKey);
        } else {
            return ((FormRow) rows.get(x)).get("id").equals(rowKey);
        }
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
        if (parent == null) parent = ".";

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
