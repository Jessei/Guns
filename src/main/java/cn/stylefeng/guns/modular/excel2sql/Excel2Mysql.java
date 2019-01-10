package cn.stylefeng.guns.modular.excel2sql;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * Copyright © 2019 Westrip Info. Tech Ltd. All rights reserved.
 *
 * @Company www.wetrip.fun
 * @Author bigWang
 * @Email wang.li@westrip.com
 * @Date 2019-01-10 20:33
 * @Version 1.0
 */

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Excel2Mysql {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    //判断Excel的版本,获取Workbook
    public static Workbook getWorkbok(InputStream in,File file) throws IOException{
        Workbook wb = null;
        if(file.getName().endsWith(EXCEL_XLS)){  //Excel 2003
            wb = new HSSFWorkbook(in);
        }
//        else if(file.getName().endsWith(EXCEL_XLSX)){  // Excel 2007/2010
//            wb = new XSSFWorkbook(in);
//        }
        return wb;
    }

    //判断文件是否是excel
    public static void checkExcelVaild(File file) throws Exception{
        if(!file.exists()){
            throw new Exception("文件不存在");
        }
        if(!(file.isFile() && (file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))){
            throw new Exception("文件不是Excel");
        }
    }

    //由指定的Sheet导出至List
    public static void exportListFromExcel() throws IOException {

        SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
        BufferedWriter bw = new BufferedWriter(new FileWriter(new File("d:/xxx/insertSql.txt")));
        try {
            // 同时支持Excel 2003、2007
            File excelFile = new File("D:/price_basic_rule.xls"); // 创建文件对象
            FileInputStream is = new FileInputStream(excelFile); // 文件流
            checkExcelVaild(excelFile);
            Workbook workbook = getWorkbok(is,excelFile);
            //Workbook workbook = WorkbookFactory.create(is); // 这种方式 Excel2003/2007/2010都是可以处理的

            int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量
            /**
             * 设置当前excel中sheet的下标：0开始
             */
            Sheet sheet = workbook.getSheetAt(0);   // 遍历第一个Sheet


            //创建表
            Date date = new Date();
            StringBuilder sb = new StringBuilder("CREATE TABLE test_");
            sb.append(fmt.format(date))
                    .append("(\n");

            // 为跳过第一行目录设置count
            int count = 0;

            for (Row row : sheet) {
//                // 跳过第一行的目录
                if(count == 0){



//                    count++;
//                    continue;
                }
                String tableValue ="";
                // 如果当前行没有数据，跳出循环
                if(row.getCell(0).toString().equals("")){
                    return ;
                }
                String rowValue = "";
                for (Cell cell : row) {
                    if(cell.toString() == null){
                        continue;
                    }
                    int cellType = cell.getCellType();
                    String cellValue = "";
                    switch (cellType) {
                        case Cell.CELL_TYPE_STRING:     // 文本
                            cellValue = "'"+cell.getRichStringCellValue().getString() + "'#";
                            tableValue = "AA"+String.valueOf(cell.getColumnIndex()) + " varchar(2000) ,";
                            break;
                        case Cell.CELL_TYPE_NUMERIC:    // 数字、日期
                            if (DateUtil.isCellDateFormatted(cell)) {
                                cellValue = fmt.format(cell.getDateCellValue()) + "#";
                                tableValue = "AA"+String.valueOf(cell.getColumnIndex()) + " varchar(2000) ,";
                            } else {
                                cell.setCellType(Cell.CELL_TYPE_STRING);
                                cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#";
                                tableValue = "AA"+String.valueOf(cell.getColumnIndex()) + " varchar(2000) ,";

                            }
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:    // 布尔型
                            cellValue = String.valueOf(cell.getBooleanCellValue()) + "#";
                            tableValue = "AA"+String.valueOf(cell.getColumnIndex()) + " varchar(2000) ,";
                            break;
                        case Cell.CELL_TYPE_BLANK: // 空白
                            cellValue = cell.getStringCellValue() + "#";
                            tableValue = "AA"+String.valueOf(cell.getColumnIndex()) + " varchar(2000) ,";

                            break;
                        case Cell.CELL_TYPE_ERROR: // 错误
                            cellValue = "错误#";
                            tableValue = "AA"+String.valueOf(cell.getColumnIndex()) + " varchar(2000) ,";

                            break;
                        case Cell.CELL_TYPE_FORMULA:    // 公式
                            // 得到对应单元格的公式
                            //cellValue = cell.getCellFormula() + "#";
                            // 得到对应单元格的字符串
                            cell.setCellType(Cell.CELL_TYPE_STRING);
                            cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#";
                            tableValue = "AA"+String.valueOf(cell.getColumnIndex()) + " varchar(2000) ,";

                            break;
                        default:
                            cellValue = "#";
                            tableValue = "AA"+String.valueOf(cell.getColumnIndex()) + " varchar(2000) ,";

                    }
                    sb.append(tableValue);
                    //System.out.print(cellValue);
                    rowValue += cellValue;
                }
                if(count == 0) {
                    count++;
                    String substring = sb.substring(0, sb.length() - 1);
                    substring += " );";
                    System.out.println(substring);
                }
                writeSql(rowValue,bw);
//                System.out.println(rowValue);
//                System.out.println();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally{
            bw.flush(); // 把缓存区内容压入文件
            bw.close(); // 最后记得关闭文件
        }
    }




    // 将值拼成sql语句
    public static void writeSql(String rowValue, BufferedWriter bw ) throws IOException {
        String[] sqlValue = rowValue.split("#");
        String sqlName = " INSERT INTO table_name (";
        String sqlValue1 = " VALUES(";
        for (int i = 0; i < sqlValue.length; i++) {
            sqlName += "AA"+(i)+"," ;
            sqlValue1 += sqlValue[i].trim() + "," ;
        }
        String substring = sqlName.substring(0, sqlName.length() - 1);
        substring +=" )";

        String substring1 = sqlValue1.substring(0, sqlValue1.length() - 1);
        substring1 += " ) ;";

        String sql = substring + substring1;
        System.out.print(sql);

        try {
            bw.write(sql);
            bw.newLine();


        } catch (IOException e) {
            e.printStackTrace();
        }

    }
    public static void main(String[] args) {

        try {
            exportListFromExcel();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
