import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;


import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * RunWin.java
 * Created at 2013-09-11
 * Created by hanwei
 * Copyright (C) 2013 SHANGHAI VOLKSWAGEN, All rights reserved.
 */

/**
 * <p>ClassName: RunWin</p>
 * <p>Description: TODO</p>
 * <p>Author: hanwei</p>
 * <p>Date: 2013-9-13</p>
 */
public class RunWin {
    public static void main(String[] args) throws Exception {
        RunWin runwin=new RunWin();
        String path="d:/order.xlsx";
        runwin.waitDueFile=path;
        runwin.setWorkbook(new File(path));
        HashMap map=runwin.parseExcel();
        String data[][]=runwin.procData((String [][])map.get("data"));
        
        runwin.setStartRowNum(0);
        runwin.setDataArray(data);
        runwin.createExcel();
        
        System.out.println("success");
    }
    
    String waitDueFile;
    /**
     * <p>
     * Field dataList: 数据list
     * </p>
     */
    private String[][] dataArray;

    /**
     * <p>
     * Field dataList: 数据list
     * </p>
     */
    private List dataList;

    /**
     * <p>Field dataListDissimilar: 任意长度数据List</p>
     */
    private List<List<Object>> dataListDissimilar;
    /**
     * <p>
     * Field startRowNum: 从第行开始写入数据
     * </p>
     */
    private int startRowNum;
    /**
     * <p>
     * Field condition: 条件
     * </p>
     */
    private String condition;
    /**
     * <p>
     * Field fileName: 生成文件的名称
     * </p>
     */
    private String fileName;
    /**
     * <p>
     * Field templatePath: 模板的路径
     * </p>
     */
    private String templatePath;

    /**
     * <p>
     * Field columns: 列名数组
     * </p>
     */
    private String[] columns;
    
    /**
     * <p>
     * Field sheetNum: sheet的序号，从0开始
     * </p>
     */
    private int sheetNum;

    /**
     * <p>
     * Field encoding: 编码
     * </p>
     */
    private String encoding = "UTF-8";
    XSSFWorkbook workbook;
    /**
     * <p>
     * Description: 创建excel
     * </p>
     * 
     * @param response response对象
     * @throws Exception 异常
     */
    public void createExcel() throws Exception {
        InputStream inputStream = null;
        OutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(new File(waitDueFile));
            outputStream = new FileOutputStream(new File("d:/output.xlsx"));

            workbook = new XSSFWorkbook(inputStream);
            XSSFSheet  sheet = workbook.getSheetAt(this.sheetNum);
            if (!"".equals(this.condition)) {
                XSSFRow row = sheet.createRow(0);
                row.createCell(0).setCellValue(this.condition);
            }
            createDataFromArray(sheet);
            workbook.write(outputStream);
        } catch (IOException e) {
            new Exception(e);
        } finally {
            try {
                outputStream.close();
                inputStream.close();
            } catch (IOException ex) {
                new Exception(ex);
            }
        }
    }

    /**
     * <p>
     * Description: 从数据数组生成文件
     * </p>
     * 
     * @param sheet excel的sheet
     */
    private void createDataFromArray(XSSFSheet sheet) {
        XSSFColor cr=new XSSFColor(new Color(255,153,204));
        XSSFColor cy=new XSSFColor(Color.yellow);
        
        XSSFCellStyle styleRed = workbook.createCellStyle();//新建样式对象  
        //style.setWrapText(true);//设置自动换行  
        styleRed.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setFillBackgroundColor(c);
        styleRed.setFillForegroundColor(cr);
        XSSFFont font = (XSSFFont) workbook.createFont();//创建字体对象  
        font.setFontName("微软雅黑");  
        //font.setColor(HSSFColor.RED.index);
        //style.setFont(font); 
        
        XSSFCellStyle styleYellow = workbook.createCellStyle();//新建样式对象  
        //style.setWrapText(true);//设置自动换行  
        styleYellow.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setFillBackgroundColor(c);
        styleYellow.setFillForegroundColor(cy);
        
        boolean curr_red=true;
        
        for (int i = 0; i < this.dataArray.length; i++) {
            Row row = sheet.createRow(i + this.startRowNum);
            if(null==dataArray[i][0] || "".equals(dataArray[i][0])){
                continue;
            }
            for (int j = 0; j < this.dataArray[i].length; j++) {
                Cell cell=row.createCell(j);
                if(i!=0 && i!=this.dataArray.length-1 && (dataArray[i][0].equals(dataArray[i+1][0]) || dataArray[i][0].equals(dataArray[i-1][0]) )){
                    if(curr_red){
                        cell.setCellStyle(styleRed);
                    }else{
                        cell.setCellStyle(styleYellow);
                    }
                    if(!dataArray[i][0].equals(dataArray[i+1][0]) && j==this.dataArray[i].length-1){
                        if(curr_red){
                            curr_red=false;
                        }else{
                            curr_red=true;;
                        }
                    }
                }else if(i==this.dataArray.length-1 && dataArray[i][0].equals(dataArray[i-1][0])){
                    if(curr_red){
                        cell.setCellStyle(styleRed);
                    }else{
                        cell.setCellStyle(styleYellow);
                    }
                }
           
                if (i != 0
                        && (j == 3 || j == 4 || j == 5 || j == 6 || j == 7 || j == 8 || j == 9 || j == 20 || j == 24
                                || j == 25 || j == 28)) {
                    String value=this.dataArray[i][j];
                    if(null==value || "".equals(value)){
                        
                    }else{
                        cell.setCellValue(Double.valueOf(value));
                    }
                }else if(i!=0 && (j==16)){
                    String value=this.dataArray[i][j];
                    if(null!= value && !"".equals(value)){
                        value=value.substring(1);
                    }
                    cell.setCellValue(value);
                }else{
                    cell.setCellValue(this.dataArray[i][j]);
                }
            }
        }
    }

    /**
     * <p>Description: 从数据数组生成文件</p>
     * @param sheet excel的sheet
     */
    private void createDataFromListDissimilar(XSSFSheet sheet) {
        if (null != this.dataListDissimilar) {
            for (int i = 0; i < this.dataListDissimilar.size(); i++) {
                XSSFRow row = sheet.createRow(i + this.startRowNum);
                if (null != this.dataListDissimilar.get(i)) {
                    for (int j = 0; j < this.dataListDissimilar.get(i).size(); j++) {
                        
                        Object value = this.dataListDissimilar.get(i).get(j);
                        String text = "";
                        if (null != value) {
                            text = value.toString();
                            row.createCell(j).setCellValue(text);
                        } 
                        
                    }
                } 
                
            }
        } 
    }

    /**
     * <p>
     * Description: 转换编码
     * </p>
     * 
     * @param value 需要转换编码的字符串
     * @return String 转换后的字符串
     */
    public String changeEncoding(String value) {
        String val = "";
        try {
           // val = new String(value.getBytes() , this.encoding);
            val = URLEncoder.encode(value, this.encoding);
            if (value.endsWith("+")) {
                value = value.substring(0, value.length() - 1);
            }
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return val;
    }

    public int getStartRowNum() {
        return this.startRowNum;
    }

    public void setStartRowNum(int startRowNum) {
        this.startRowNum = startRowNum;
    }

    public String getCondition() {
        return this.condition;
    }

    public void setCondition(String condition) {
        this.condition = condition;
    }

    public String getFileName() {
        return this.fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getTemplatePath() {
        return this.templatePath;
    }

    public void setTemplatePath(String templatePath) {
        this.templatePath = templatePath;
    }

    public int getSheetNum() {
        return this.sheetNum;
    }

    public void setSheetNum(int sheetNum) {
        this.sheetNum = sheetNum;
    }

    public String getEncoding() {
        return this.encoding;
    }

    public void setEncoding(String encoding) {
        this.encoding = encoding;
    }

   

    public List getDataList() {
        return this.dataList;
    }

    public String[][] getDataArray() {
        return this.dataArray;
    }

    public void setDataArray(String[][] dataArray) {
        this.dataArray = dataArray;
    }

    public String[] getColumns() {
        return this.columns;
    }

    public void setColumns(String[] columns) {
        this.columns = columns;
    }

    public void setDataList(List dataList) {
        this.dataList = dataList;
    }

    public List<List<Object>> getDataListDissimilar() {
        return this.dataListDissimilar;
    }

    public void setDataListDissimilar(List<List<Object>> dataListDissimilar) {
        this.dataListDissimilar = dataListDissimilar;
    }

    
    
    
    

    /**
     * <p>Field MAX_PARSABLE_ROWS: 可解析的最大行数</p>
     */
    public static final int MAX_PARSABLE_ROWS = 100000;
    /**
     * <p>Field wb: workbook对象</p>
     */
    private XSSFWorkbook wb = null;



    
    
    
    /**
     * Workbook setter method, change the workbook object at runtime 
     * using InputStream object.
     * It provents the situation that the position is at the end of
     * this stream.
     * @param tis   InputStream object where workbook object generate from.
     * @throws Exception The unchecked exception
     */
    public void setWorkbook(InputStream tis) throws Exception {
        try {
            if(tis.available() == 0){
                throw new Exception("文件不存在或已至文件流末尾");
            }
        }catch(IOException exc) {
            throw new Exception("输入输出失败");
        }
        getWorkbookFromStream(tis);
    }
    
    /**
     * Workbook setter method, change the workbook object at runtime 
     * using File object.
     * @param tfile File object where workbook object generate from.
     * @throws Exception The unchecked exception
     */
    public void setWorkbook(File tfile) throws Exception {
        if (!isExcelFile(tfile)) {
            throw new Exception("请确认文件类型");
        }
        getWorkbookFromFile(tfile);
    }
    
    /**
     * Generate Workbook object from InputStream object. 
     * Change the workbook object dynamically, 
     * close the previous Workbook object.
     * @param tis   InputStream object.
     * @throws Exception The unchecked exception
     */
    private void getWorkbookFromStream(InputStream tis) throws Exception {
        try {
            if(wb == null){
                wb=new XSSFWorkbook(tis);
            }
        } catch (IOException bexc) {
            throw new Exception(
                    "获得EXCEL文件输入流失败");
        } 
    }
    
    /**
     * Generate Workbook object from File object.
     * Change the workbook object dynamically, 
     * close the previous Workbook object.
     * @param tfile File object.
     * @throws Exception The unchecked exception
     */
    private void getWorkbookFromFile(File tfile) throws Exception {
        InputStream tis;
        try {
            tis = new FileInputStream(tfile);
        } catch (FileNotFoundException e) {
            throw new Exception(e);
        }
        getWorkbookFromStream(tis);
    }
    
    
    /**
     * Parse and get data from the workbook Get all the data of sheet 0.
     * 
     * @return HashMap The result of the process <br>
     *         <strong>The Structure Of The HashMap:</strong> <br>
     *         <b><i>key</i></b>        <b>value</b><br>
     *         <i>nextrow</i>----An Integer object indicates the next row's number 
     *                           which haven't been parsed.
     *                           -1 means the end of the file <br>
     *         <i>data</i>-------A two dimension String array contains 
     *                           the data of the file 
     *                           which area have already been parsed.
     *                           It can be NULL if no data found.<br>
     * @see getExcelData(Workbook twb,int sheetno,int rows)
     * @throws Exception The unchecked exception
     */
    public HashMap parseExcel() throws Exception {
        return parseExcel(0);
    }

    /**
     * Parse and get data from the file; Get all the data of specified sheet.
     * 
     * @param sheetno
     *            The index of the sheet which will be parsed. Begins from 0.
     * @return  HashMap 数据map对象
     * @throws Exception 异常
     */
    public HashMap parseExcel(int sheetno)
            throws Exception {
        if (sheetno >= wb.getNumberOfSheets() || sheetno < 0) {
            throw new Exception("没有指定的Sheet");
        }
        XSSFSheet sheet = wb.getSheetAt(sheetno);
        if (sheet.getLastRowNum() > MAX_PARSABLE_ROWS) {
            throw new Exception("单次返回太多行");
        }
        return getExcelData(sheetno, 0, sheet.getLastRowNum());
    }

    /**
     * Parse and get data from the workbook.
     * Get specified rows of data of specified
     * sheet,specified row index.
     * 
     * @param sheetno
     *            The index of the sheet which will be parsed.
     * @param currow
     *            The index of the row where parse process start from. Begin from 0. 
     * @param rows
     *            The number of the rows to be parsed.
     * @return  HashMap 数据map对象
     * @throws Exception The unchecked exception
     */
    public HashMap parseExcel(int sheetno, int currow, int rows) 
    throws Exception {
        if (rows > MAX_PARSABLE_ROWS) {
            throw new Exception("单次返回太多行");
        }
        if (currow < 0) {
            throw new Exception("数据解析起始行索引无效");
        }
        if (rows <= 0) {
            throw new Exception("获取数据行的行数无效");
        }
        
        if (sheetno >= wb.getNumberOfSheets() || sheetno < 0) {
            throw new Exception("没有指定的Sheet");
        }
        return getExcelData(sheetno, currow, rows);
    }

    /**
     * Check the parameter tFile is a Excel-Format File The check process just
     * depends on the suffix of the file's name.
     * 
     * @param tFile
     *            The file to be checked.
     * @return true The file is a excel file false The file isn't a excel file
     */
    private boolean isExcelFile(File tFile) {
        if (tFile.getName().toLowerCase().endsWith("xlsx")) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * 获取excel的数据
     * @param sheetno
     *            The index of the sheet which will be parsed; begin from 0.
     * @param currow
     *            The index of the row where parse process start from.
     * @param rows
     *            The number indicates how many rows will be parsed.
     * @return HashMap The result of the parsed excel file. <br>
     *         <strong>The Structure Of The HashMap:</strong> <br>
     *         <b><i>key</i></b>        <b>value</b><br>
     *         <i>nextrow</i>----An Integer object indicates the next row's number 
     *                           which haven't been parsed
     *                           -1 means the end of the file. <br>
     *         <i>data</i>-------A two dimension String array contains 
     *                           the data of the file 
     *                           which area have already been parsed. 
     *                           It can be NULL if no data found.<br>
     * @throws Exception 异常
     */
    /**
     * <p>Description: TODO</p>
     * @param sheetno
     * @param currow
     * @param rows
     * @return
     * @throws Exception
     */
    /**
     * <p>Description: TODO</p>
     * @param sheetno
     * @param currow
     * @param rows
     * @return
     * @throws Exception
     */
    private HashMap getExcelData(int sheetno, int currow, int rows)
            throws Exception {
        HashMap hm = new HashMap();
        XSSFSheet sheet = wb.getSheetAt(sheetno);
        int rowlmt = 0;
        /*
         * Define variable rowlmt to get the number of the rows that will be
         * parsed exactly
         */
        if (currow + rows >= sheet.getLastRowNum()) {
            rowlmt = sheet.getLastRowNum() ;
            hm.put("nextrow", new Integer(-1));
        } else {
            rowlmt = currow + rows ;
            hm.put("nextrow", new Integer(rowlmt + 1));
        }
        String[][] rst = null;
        int columnNum=sheet.getRow(sheet.getFirstRowNum()).getLastCellNum();
        if(rowlmt - currow + 1 > 0) {
            rst = new String[rowlmt - currow + 1][columnNum];
        }
        /*
         * Put data into a two dimension String array. The value of the cell
         * depends on its celltype
         */
        for (int i = currow; i <= rowlmt; i++) {
            for (int j = 0; j <columnNum; j++) {
                try {
                    XSSFCell cell = sheet.getRow(i).getCell(j);
                    if(cell==null){
                        rst[i - currow][j] = "";
                        continue;
                    }
                    
                    if (cell.getCellType()==XSSFCell.CELL_TYPE_NUMERIC) {
                        if(HSSFDateUtil.isCellDateFormatted(cell)){
                            SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd HH:mm");
                            Date date = cell.getDateCellValue();
                            String val=sdf.format(date);
                            rst[i - currow][j] = val;
                        }else{
                            DecimalFormat df = new DecimalFormat("###############.########");
                            String val = df.format(cell.getNumericCellValue());
                            rst[i - currow][j] = val;
                        }
                    }else if (cell.getCellType()==XSSFCell.CELL_TYPE_BLANK) {
                        rst[i - currow][j] = "";
                    }else if (cell.getCellType()==XSSFCell.CELL_TYPE_FORMULA) {
                        DecimalFormat df = new DecimalFormat("###############.########");
                        String val="";
                        if (cell.getCachedFormulaResultType()==XSSFCell.CELL_TYPE_NUMERIC) {
                            double nfcv = cell.getNumericCellValue();
                            val = df.format(nfcv);
                        }else if(cell.getCachedFormulaResultType()==XSSFCell.CELL_TYPE_STRING){
                            val = cell.getStringCellValue();
                        }else if(cell.getCachedFormulaResultType()==XSSFCell.CELL_TYPE_BLANK){
                           val= "";
                        }
                        rst[i - currow][j] = val;
                    }else if (cell.getCellType()==XSSFCell.CELL_TYPE_STRING) {
                        rst[i - currow][j] = cell.getStringCellValue();
                    }else {
                        rst[i - currow][j] = "";
                    }
                } catch (Exception exc) {
                    throw new Exception("EXCEL文件单元格解析失败，请检查文件格式");
                }
            }
        }
        hm.put("data", rst);
        return hm;
    }
    
    
    public void show(String [][]data){
        for(int i=0;i<data.length;i++){
            for(int j=0;j<data[i].length;j++){
                System.out.println(data[i][j]);
            }
        }
    }
    
    public String[][] procData(String [][]data){
        List<String[]> totalList=new ArrayList();
        for(int i=0;i<data.length;i++){
            String [] tmpArray=new String[37];
            for(int m=0;m<19;m++){
                tmpArray[m]=data[i][m];
            }
            for(int m=20;m<37;m++){
                tmpArray[m]=data[i][m];
            }
            String title=data[i][19];
            if(title.indexOf("，")!=-1){
                String [] titles=title.split("，");
                for(int k=0;k<titles.length;k++){
                    String [] tmp2=new String[37];
                    copyArray(tmp2, tmpArray);
                    tmp2[19]=titles[k];
                    totalList.add(tmp2);
                }
            }else{
                tmpArray[19]=data[i][19];
                totalList.add(tmpArray);
            }
        }
        int len=totalList.size();
        String [][] datas=new String [len][37];
        
        for(int n=0;n<len;n++){
            Object[] tl=totalList.get(n);
            for(int j=0;j<37;j++){
                datas[n][j]=String.valueOf(tl[j]);
            }
        }
        return datas;
    }
    
    private void copyArray(String[]d1,String[]d2){
        for(int i=0;i<d2.length;i++){
            d1[i]=d2[i];
        }
    }
}
