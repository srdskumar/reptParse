/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapp;

/**
 *
 * @author g706134
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
 

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class SimpleExcelReaderExample {
     
    public static void main(String[] args) throws IOException {
        String excelFilePath = "Books.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
         
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet uninv_open = workbook.getSheetAt(5);
        Iterator<Row> iterator = uninv_open.iterator();
         
                String rec = "";
                String pay = "";
                String per = "";
                String svc = "";
                double dval = 0;
                double cval = 0;
                
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
             dval=0;
             cval=0;
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell.getColumnIndex() == 1 || cell.getColumnIndex() == 2 || cell.getColumnIndex() == 5 || cell.getColumnIndex() == 8 || cell.getColumnIndex() == 45 || cell.getColumnIndex() == 46)  
                {
                
                  
                        
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        
                        if(cell.getColumnIndex() == 1)
                        {
                            rec=cell.getStringCellValue();
                        }
                        if (cell.getColumnIndex() == 2 )
                        {
                            pay=cell.getStringCellValue();
                        }
                        if (cell.getColumnIndex() == 5 )
                        {
                            svc=cell.getStringCellValue();
                        }
                         if (cell.getColumnIndex() == 8 )
                        {
                            per=cell.getStringCellValue();
                        }
                         
                                                 
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        //System.out.print(cell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        //System.out.print(cell.getNumericCellValue());
                        if (cell.getColumnIndex() == 45 )
                        {
                            dval=cell.getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == 46 )
                        {
                            cval=cell.getNumericCellValue();
                        }
                        break;
                }
                
                }
                
            }
            System.out.print(rec+"-"+pay+"-"+per+"-"+svc+"|"+dval+"|"+cval);
            System.out.println();
        }
         
        workbook.close();
        inputStream.close();
    }
 
}





