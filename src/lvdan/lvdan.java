package lvdan;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class lvdan {
    public static void main(String[] args){
        //读取Excel表格
        readXml("C:\\Users\\sszz\\Desktop\\赤坂亭\\8.6\\工作簿1.xls");
    }
    public static void readXml(String fileName){
        boolean isE2007 = false; //判断是否是Excel2007格式
        if(fileName.endsWith("xlsx"))
            isE2007 = true;
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
            while(rows.hasNext()){
                Row row = rows.next();//获得行数据；
                int colNum = row.getPhysicalNumberOfCells();
               // System.out.println("第"+colNum+"列");
                Iterator<Cell> cells = row.cellIterator();//获得第一行的迭代器

                while (cells.hasNext()){
                    Cell cell = cells.next();
                    Cell a = null;
                    Cell b = null;
                    Cell c = null;
                    Row row1 = null;
                    int i = 0;
                    switch (cell.getCellType()) {   //根据cell中的类型来输出数据
                        case HSSFCell.CELL_TYPE_NUMERIC:
                            //row1 = sheet.getRow(0);
                            //a = row1.getCell(cell.getColumnIndex());
                            row1 = sheet.getRow(1);
                            b = row1.getCell(cell.getColumnIndex());
                            row1 = sheet.getRow(2);
                            c = row1.getCell(cell.getColumnIndex());
                            System.out.println(b+"       "+c+"       "+cell.getNumericCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_STRING:
                            System.out.println();
                            System.out.println("店面      "+cell.getStringCellValue());
                            System.out.println("货品名称----"+"单位----"+"数量");
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN:
                            System.out.println("第"+row.getRowNum()+"行"+" "+"第"+cell.getColumnIndex()+"列"+":"+cell.getBooleanCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_FORMULA:
                            System.out.println("第"+row.getRowNum()+"行"+" "+"第"+cell.getColumnIndex()+"列"+":"+cell.getCellFormula());
                            break;
                        case HSSFCell.CELL_TYPE_BLANK:
                            break;
                    }
                }

            }
        }catch (IOException ex){
            ex.printStackTrace();
        }
    }
}
