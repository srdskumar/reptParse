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



public class ParseDebtorTransactionsData {
     
    public static void main(String[] args) throws IOException {
        
        //open(5),close(24),datacorrections-2(14,16),adjustments-2(19,21),o1cf(22),controldata(37,0),reconciledinvoice(35,36)
        
        //String sheet_names[] = {"Uninv Opening Position","Uninv Closing Position","Debtor Reconciled Invoices","Uninv Debtor Data Corrections","Uninv Debtor Adjustments","Uninv One1Clear Features","Debtor Control Data"};            
        ArrayList<String> open = parseReport("Opening Position",1,2,5,8,44,46,"open");
        ArrayList<String> close = parseReport("Closing Position",1,2,5,8,44,46,"close");
        //ArrayList<String> rinvoice = parseReport("Debtor Reconciled Invoices",0,1,2,3,5,6,"rinvoice");
        ArrayList<String> rinvoice = parseReport("Debtor Reconciled Invoices",0,1,2,3,10,10,"rinvoice");
        ArrayList<String> correction = parseReport("Debtor Data Corrections",0,1,4,5,10,10,"correction");
        ArrayList<String> adjust = parseReport("Debtor Adjustments",0,4,7,8,11,11,"adjust");
        ArrayList<String> o1cf = parseReport("One1Clear Features",1,2,5,8,41,43,"o1cf");
        ArrayList<String> settled = parseReport("Settled Transactions-Funds-Paid",0,1,5,4,19,20,"settled");
        ArrayList<String> alloc = parseReport("Cash Allocations",9,10,11,12,19,21,"alloc");
        ArrayList<String> writeoff = parseReport("Write_off",9,10,11,12,19,21,"writeoff");
        
        
        cleanDB();
        //importDB(open);
        
        Connection conn = null;
        Statement stmt = null;
        String sql = "";
                 
        try {
            
        Connection con = DBConnection.getConnection();
        stmt = con.createStatement();
        
        System.out.println("Loading data to OPEN");
        Iterator<String> itr=open.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
        
        }
        
        System.out.println("Loading data to CLOSE");
        itr=close.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
        }
         
        System.out.println("Loading data to RINVOICE");
        itr=rinvoice.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
        }
        
        System.out.println("Loading data to CORRECTION");
        itr=correction.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
        }
        
        System.out.println("Loading data to ADJUST");
        itr=adjust.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
        }
        
        System.out.println("Loading data to O1CF");
        itr=o1cf.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
       // System.out.println(sql); 
        }
        
        System.out.println("Loading data to SETTLED");
        itr=settled.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        //System.out.println(sql); 
        stmt.executeUpdate(sql);
        
        }
        
        System.out.println("Loading data to ALLOC");
        itr=alloc.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
        }
        
        System.out.println("Loading data to WRITE OFF");
        itr=writeoff.iterator();  
        while(itr.hasNext()){  
        sql = itr.next();
        stmt.executeUpdate(sql);
        //System.out.println(sql); 
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
        stmt.executeUpdate("delete from settled");
        stmt.executeUpdate("delete from alloc");
        stmt.executeUpdate("delete from writeoff");
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
                String filter = "";
                
                
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
             dval=0;
             cval=0;
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                
                
                if( cell.getColumnIndex() == 23 || cell.getColumnIndex() == c1 || cell.getColumnIndex() == c2 || cell.getColumnIndex() == c3 || cell.getColumnIndex() == c4 || cell.getColumnIndex() == c5 || cell.getColumnIndex() == c6)  
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
                         
                         if (cell.getColumnIndex() == 23 )
                        {
                            filter=cell.getStringCellValue();
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
            if(rec.length() == 5 || rec.length() == 8){
            rpps = rec+"-"+pay+"-"+per+"-"+svc;
            
           if(tbl.equalsIgnoreCase("open") || tbl.equalsIgnoreCase("close")) 
           {
               if(filter.equalsIgnoreCase("Missing Invoice") || filter.equalsIgnoreCase("Unreconciled") || filter.equalsIgnoreCase(""))
                       {
                           continue;
                       }
               
           }
            
           // if(name.equalsIgnoreCase("Settled Transactions-Funds-Paid")){
            //System.out.println(rpps+"|"+dval+"|"+cval);
            //}
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





