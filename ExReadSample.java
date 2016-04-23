/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapp;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.*;
 

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;


/**
 *
 * @author g706134
 */
public class ExReadSample {
    
    
    public static void main(String[] args) throws IOException {
        
        String excelFilePath = "GBRCNCOR.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        
        /*
        ArrayList<String> open = parseReport(5,1,2,5,8,44,46,"open");
        ArrayList<String> close = parseReport(24,1,2,5,8,44,46,"close");
        ArrayList<String> rinvoice = parseReport(34,0,1,2,3,5,6,"rinvoice");
        ArrayList<String> correction = parseReport(14,0,1,4,5,10,10,"correction");
        ArrayList<String> adjust = parseReport(18,0,4,7,8,11,11,"adjust");
        ArrayList<String> o1cf = parseReport(22,1,2,8,5,44,46,"o1cf");
        ArrayList<String> cdata = parseReport(36,0,1,2,3,7,7,"cdata");
        */
        
        //ArrayList<String> sheet_names = {"Uninv Opening Position","Uninv Closing Position","Debtor Reconciled Invoices","Uninv Debtor Data Corrections","Uninv Debtor Adjustments","Uninv One1Clear Features","Debtor Control Data"};
        String sheet_names[] = {"Uninv Opening Position","Uninv Closing Position","Debtor Reconciled Invoices","Uninv Debtor Data Corrections","Uninv Debtor Adjustments","Uninv One1Clear Features","Debtor Control Data"};
        int sheet_no;
        for ( String str : sheet_names ) 
        {
            sheet_no = workbook.getSheetIndex(str);
            Sheet wb_sheet = workbook.getSheetAt(sheet_no);
            
            String sheet_name = str;
        Iterator<Row> iterator = wb_sheet.iterator();
        
                String rec = "";
                String pay = "";
                String per = "";
                String svc = "";
                double dval = 0;
                double cval = 0;
                String rpps = "";
                int j=0;
                
                while (iterator.hasNext()) {
                    j++;
                    
                    System.out.println(sheet_name+"----->row"+j);
                    if(j == 10){j=0;break;}
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
             dval=0;
             cval=0;
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell.getColumnIndex() == 1 || cell.getColumnIndex() == 2 || cell.getColumnIndex() == 5 || cell.getColumnIndex() == 8 || cell.getColumnIndex() == 44 || cell.getColumnIndex() == 46)  
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
                        if (cell.getColumnIndex() == 44 )
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
            if(rec.length() == 5 || rec.length() == 8){
            rpps = rec+"-"+pay+"-"+per+"-"+svc;
            System.out.println(rpps+"|"+dval+"|"+cval);
            //System.out.println("insert into "+tbl+" (rpps,sdrval) values(\""+rpps+"\","+dval+")");
            //System.out.println();
            // ADD insert Query to Array 
            //stmt.executeUpdate("insert into "+tbl+" (rpps,sdrval) values(\""+rpps+"\","+dval+")");
            }
        }
        
        }
        
        
        
       
        
        
        
    }
    
}
