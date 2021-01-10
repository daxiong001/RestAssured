package excelutils;

import com.sun.org.apache.bcel.internal.util.ClassPath;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class ExcelReader {

    private static Logger logger = Logger.getLogger(ExcelReader.class.getName());

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";


    /**
     * 判断excel文件类型创建具体对象
     *
     * @param inputStream
     * @param fileType
     * @return
     * @throws IOException
     */
    public static Workbook getWorkBook(InputStream inputStream, String fileType) throws IOException {
        Workbook workbook = null;
        if (fileType.equalsIgnoreCase(XLS)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (fileType.equalsIgnoreCase(XLSX)) {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }


    public static List<ExcelModel> readExcel(String fileName) {
        Workbook workbook = null;
        FileInputStream inputStream = null;
        try {
            //获取文件名后缀
            String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
            File excelFile = new File(fileName);
            if (!excelFile.exists()) {
                try {
                    throw new Exception("指定的Excel文件不存在！");
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            //获取excel工作薄对象

            inputStream = new FileInputStream(excelFile);
            workbook = getWorkBook(inputStream, fileType);

            List<ExcelModel> resultDataList = parseExcel(workbook);
            return resultDataList;
        }catch (Exception e){
            logger.warning("解析Excel失败，文件名：" + fileName + " 错误信息：" + e.getMessage());
            return null;
        } finally {
            try {
                if (null != workbook) {
                    workbook.close();
                }
                if (null != inputStream) {
                    inputStream.close();
                }
            } catch (Exception e) {
                logger.warning("关闭数据流出错！错误信息：" + e.getMessage());
                return null;
            }
        }
    }





    private static List<ExcelModel> parseExcel(Workbook workbook){
        List<ExcelModel> resultDataList = new ArrayList<ExcelModel>();
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
            Sheet sheet = workbook.getSheetAt(sheetNum);
            if (sheet == null){
                continue;
            }
            int firstRowNum = sheet.getFirstRowNum();
            Row firstRow = sheet.getRow(firstRowNum);
            if (null == firstRow){
                logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
            }
            int rowStart = firstRowNum + 1;
            int rowEnd = sheet.getPhysicalNumberOfRows();
            for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null){
                    continue;
                }
                ExcelModel resultData = convertRowToData(row);
                if (resultData == null){
                    logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
                    continue;
                }
                resultDataList.add(resultData);
            }
        }

        return resultDataList;
    }


    private static String convertCellValueToString(Cell cell){
        if(cell==null){
            return null;
        }
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:   //数字
                Double doubleValue = cell.getNumericCellValue();

                // 格式化科学计数法，取一位整数
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                break;
            case FORMULA:   // 公式
                returnValue = cell.getCellFormula();
                break;
            case ERROR:     // 故障
                break;
            default:
                break;
        }
        return returnValue;
    }

    private static ExcelModel convertRowToData(Row row){
        ExcelModel excelModel = new ExcelModel();

        Cell cell;
        int cellNum =0;

        cell = row.getCell(cellNum++);
        String requestName = convertCellValueToString(cell);
        excelModel.setRequestName(requestName);

        cell = row.getCell(cellNum++);
        String requestMethod = convertCellValueToString(cell);
        excelModel.setRequestMethod(requestMethod);

        cell = row.getCell(cellNum++);
        String requestHeader = convertCellValueToString(cell);
        excelModel.setRequestHeader(requestHeader);

        cell = row.getCell(cellNum++);
        String requestUrl = convertCellValueToString(cell);
        excelModel.setRequestUrl(requestUrl);

        cell = row.getCell(cellNum++);
        String requestBody = convertCellValueToString(cell);
        excelModel.setRequestBody(requestBody);

        return excelModel;
    }

    /**
     * 通过请求方法名遍历Excle表格对象list获取当前请求对象
     * @param excelModel
     * @param requestName
     * @return
     */
    public static List<String> getKey(List<ExcelModel> excelModel, String requestName){
        List<String> arrayList = new ArrayList<String>();

        for (ExcelModel obj:excelModel
             ) {
            if (requestName.equals(obj.getRequestName())){
                arrayList.add(obj.getRequestName());
                arrayList.add(obj.getRequestMethod());
                arrayList.add(obj.getRequestHeader());
                arrayList.add(obj.getRequestUrl());
                arrayList.add(obj.getRequestBody());
            }
        }

        return arrayList;
    }

    public static String listParseString(List<String> list){
        StringBuilder builder = new StringBuilder();
        for (String value:list
             ) {
            builder.append(value);
        }
        return String.valueOf(builder);
    }

    public static void main(String[] args) {
        List<ExcelModel> excelObj =ExcelReader.readExcel("..\\RestAssured\\src\\main\\resources\\demo.xlsx");
        System.out.println(ExcelReader.listParseString(ExcelReader.getKey(excelObj,"q")));;

    }
}
