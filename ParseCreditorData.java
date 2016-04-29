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



public class ParseCreditorData {
     
    public static void main(String[] args) throws IOException {
        
        //open(5),close(24),datacorrections-2(14,16),adjustments-2(19,21),o1cf(22),controldata(37,0),reconciledinvoice(35,36)
        
        //String sheet_names[] = {"Uninv Opening Position","Uninv Closing Position","Debtor Reconciled Invoices","Uninv Debtor Data Corrections","Uninv Debtor Adjustments","Uninv One1Clear Features","Debtor Control Data"};            
        ArrayList<String> open = parseReport("Uninv Opening Position",1,2,5,8,46,46,"open");
        ArrayList<String> close = parseReport("Uninv Closing Position",1,2,5,8,46,46,"close");
        ArrayList<String> rinvoice = parseReport("Creditor Reconciled Invoices",0,1,2,3,5,6,"rinvoice");
        ArrayList<String> correction = parseReport("Uninv Creditor Data Corrections",1,0,4,5,10,10,"correction");
        ArrayList<String> adjust = parseReport("Uninv Creditor Adjustments",4,0,7,8,11,11,"adjust");
        ArrayList<String> o1cf = parseReport("Uninv One1Clear Features",1,2,8,5,46,46,"o1cf");
        ArrayList<String> cdata = parseReport("Creditor Control Data",0,1,2,3,7,7,"cdata");
        
        
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
        //System.out.println(sql); 
        
        }
        
        itr=close.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
        }
         
        itr=rinvoice.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
        }
        
        itr=correction.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        //System.out.println(sql); 
        stmt.executeUpdate(sql);
         
        }
        
        itr=adjust.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        //System.out.println(sql); 
        stmt.executeUpdate(sql);
        
        }
        
        itr=o1cf.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
       // System.out.println(sql); 
        }
        
        itr=cdata.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        System.out.println(sql); 
        }
        
            
        }
        catch (SQLException se){
            se.printStackTrace();
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
        stmt.executeUpdate("delete from rinvoice");
        stmt.executeUpdate("delete from correction");    
        stmt.executeUpdate("delete from adjust");    
        stmt.executeUpdate("delete from o1cf");
        stmt.executeUpdate("delete from cdata");
        stmt.executeUpdate("delete from sanity");
        stmt.executeUpdate("delete from anomoly");
        
        }
        catch (SQLException se){
            se.printStackTrace();
        }
        
        
    }
 
    public static ArrayList<String> parseReport(String name,int c1,int c2,int c3,int c4,int c5,int c6,String tbl) throws IOException
    {
        
                    
        String excelFilePath = "S9/GBRCNCOR.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        
        ArrayList<String> ar = new ArrayList<String>();
        String sqlstr = "";
        
        Workbook workbook = new XSSFWorkbook(inputStream);
        //Sheet uninv_open = workbook.getSheetAt(sno);  
          Sheet uninv_open = workbook.getSheet(name);
        //String sname = workbook.getSheetName(sno);
        System.out.println("Parsing Sheet Name : "+name+" ---> "+tbl);
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
            if((rec.length() == 5 || rec.length() == 8) && (pay.length() == 5 || pay.length() == 8) ){
            rpps = rec+"-"+pay+"-"+per+"-"+svc;
            //System.out.println(rpps+"|"+dval+"|"+cval);
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





