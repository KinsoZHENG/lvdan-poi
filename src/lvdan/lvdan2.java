package lvdan;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;


public class lvdan2 {
    public static void main(String[] args){
        //读取Excel表格
        readXml("C:\\Users\\sszz\\Desktop\\赤坂亭\\12.30\\工作簿1.xls");
    }
    public static void readXml(String fileName) {
        boolean isE2007 = false; //判断是否是Excel2007格式
        if (fileName.endsWith("xlsx"))
            isE2007 = true;
        //定义一个新工作簿
        Workbook wb2 = new HSSFWorkbook();
        //创建sheet
        Sheet sheetNew = wb2.createSheet("sheet1");
        try {
            InputStream input = new FileInputStream(fileName); //建立表单输入流
            Workbook wb = null;
            //根据文件格式（2003/2007）初始化
            if (isE2007)
                wb = new XSSFWorkbook(input);
            else
                wb = new HSSFWorkbook(input);
            Sheet sheet = wb.getSheetAt(0);//获得第一个表单
            Iterator<Row> rows = sheet.rowIterator();//获得第一个表单的迭代器
            for (int l = 0; rows.hasNext(); l++) {
                while (rows.hasNext()) {
                    Row row = rows.next();//获得行数据；
                    int colNum = row.getPhysicalNumberOfCells();
                    Iterator<Cell> cells = row.cellIterator();//获得第一行的迭代器
                    int i = 0; // 行数计数器；
                    while (cells.hasNext()) {
                        Cell cell = cells.next();
                        Cell a = null;
                        Cell b = null;
                        Cell c = null;
                        Row row1 = null;
                        switch (cell.getCellType()) {   //根据cell中的类型来输出数据
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                Row row2 = sheetNew.createRow(l);
                                row1 = sheet.getRow(1);
                                b = row1.getCell(cell.getColumnIndex());
                                row1 = sheet.getRow(2);
                                c = row1.getCell(cell.getColumnIndex());
                                row2.createCell(0).setCellValue("");
                                row2.createCell(1).setCellValue(b.toString());
                                row2.createCell(2).setCellValue(c.toString());
                                row2.createCell(3).setCellValue(cell.getNumericCellValue());
                                l = l + 1;
                                break;
                            case HSSFCell.CELL_TYPE_STRING:
                                row2 = sheetNew.createRow(l);
                                row2.createCell(0).setCellValue(cell.getStringCellValue());
                                row2.createCell(1).setCellValue("品名");
                                row2.createCell(2).setCellValue("单位");
                                row2.createCell(3).setCellValue("数量");
                                l = l + 1;
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                row2 = sheetNew.createRow(l);
                                l = l + 1;
                                break;
                            case HSSFCell.CELL_TYPE_FORMULA:
                                System.out.println("第" + row.getRowNum() + "行" + " " + "第" + cell.getColumnIndex() + "列" + ":" + cell.getCellFormula());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:
                                break;
                        }
                    }

                }
                FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\sszz\\Desktop\\赤坂亭\\12.30\\12.30.xls");
                wb2.write(fileOutputStream);
                fileOutputStream.close();
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}