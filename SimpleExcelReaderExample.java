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
import java.util.*;
 

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;



public class SimpleExcelReaderExample {
     
    public static void main(String[] args) throws IOException {
        
        ArrayList<String> open = parseReport(5,1,2,5,8,44,46,"open");
        ArrayList<String> close = parseReport(24,1,2,5,8,44,46,"close");
        cleanDB();
        //importDB(open);
        
        Connection conn = null;
        Statement stmt = null;
        String sql = "";
                 
        try {
            
        Connection con = DBConnection.getConnection();
        stmt = con.createStatement();
        
        Iterator<String> itr=open.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        System.out.println(sql); 
        }
        
        itr=close.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        System.out.println(sql); 
        }
            
            
        }
        catch (SQLException se){
            se.printStackTrace();
        }
        
        
        Iterator<String> itr=open.iterator();  
        while(itr.hasNext()){  
        System.out.println(itr.next());  
        }  
        
        itr=close.iterator();  
        while(itr.hasNext()){  
        System.out.println(itr.next());  
        }  
        
    }
    
    public static void cleanDB()
    {
        
        Connection conn = null;
        Statement stmt = null;
        String sql = "";
                 
        try {
            
        Connection con = DBConnection.getConnection();
        stmt = con.createStatement();
        stmt.executeUpdate("delete from open");      
        stmt.executeUpdate("delete from close");
            
            
        }
        catch (SQLException se){
            se.printStackTrace();
        }
        
        
    }
 
    public static ArrayList<String> parseReport(int sno,int c1,int c2,int c3,int c4,int c5,int c6,String tbl) throws IOException
    {
        
                    
        String excelFilePath = "Books.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        System.out.print("Parsing Sheet Number ---> "+sno);
        System.out.println();
        ArrayList<String> ar = new ArrayList<String>();
        String sqlstr = "";
        
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet uninv_open = workbook.getSheetAt(sno);
        Iterator<Row> iterator = uninv_open.iterator();
         
                String rec = "";
                String pay = "";
                String per = "";
                String svc = "";
                double dval = 0;
                double cval = 0;
                String rpps = "";
                
                
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
             dval=0;
             cval=0;
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell.getColumnIndex() == c1 || cell.getColumnIndex() == c2 || cell.getColumnIndex() == c3 || cell.getColumnIndex() == c4 || cell.getColumnIndex() == c5 || cell.getColumnIndex() == c6)  
                {
                        
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        
                        if(cell.getColumnIndex() == c1)
                        {
                            rec=cell.getStringCellValue();
                        }
                        if (cell.getColumnIndex() == c2 )
                        {
                            pay=cell.getStringCellValue();
                        }
                        if (cell.getColumnIndex() == c3 )
                        {
                            svc=cell.getStringCellValue();
                        }
                         if (cell.getColumnIndex() == c4 )
                        {
                            per=cell.getStringCellValue();
                        }
                                                                        
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        //System.out.print(cell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        //System.out.print(cell.getNumericCellValue());
                        if (cell.getColumnIndex() == c5 )
                        {
                            dval=cell.getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == c6 )
                        {
                            cval=cell.getNumericCellValue();
                        }
                        break;
                }
                
                }
                
            }
            if(rec.length() == 5){
            rpps = rec+"-"+pay+"-"+per+"-"+svc;
            //System.out.print(rpps+"|"+dval+"|"+cval);
            //System.out.println("insert into "+tbl+" (rpps,sdrval) values(\""+rpps+"\","+dval+")");
            //System.out.println();
            // ADD insert Query to Array 
            sqlstr = "insert into "+tbl+" (rpps,sdrval) values(\""+rpps+"\","+dval+")";
            ar.add(sqlstr);
            //stmt.executeUpdate("insert into "+tbl+" (rpps,sdrval) values(\""+rpps+"\","+dval+")");
            }
        }
         
        workbook.close();
        inputStream.close();
        
        return ar;
        
    }
    
    
    
}





