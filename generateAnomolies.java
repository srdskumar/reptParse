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
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;


public class generateAnomolies {
    
    
    public static void main(String[] args) throws IOException {
        
        
         
        // getAnomolies("ctransrec","Creditor",0);
        // getAnomolies("transrec","Debtor",1);
         
        
        generateAnomolyReport();
        //generateAnomolyReport("Creditor");
        //generateAnomolyReport("Uninv_Debtor");
        //generateAnomolyReport("Uninv_creditor");
        
    }
    
    
    
    
     public static void generateAnomolyReport() throws IOException
    {
        
        
        XSSFWorkbook workbook = new XSSFWorkbook();        
        
        HashMap<String, List<String>> hmd = new HashMap<String, List<String>>();
        HashMap<String, List<String>> hmc = new HashMap<String, List<String>>();
        HashMap<String, List<String>> uhmd = new HashMap<String, List<String>>();
        HashMap<String, List<String>> uhmc = new HashMap<String, List<String>>();
        
        hmd = getAnomolies("transrec","Debtors",0);       
        hmc = getAnomolies("ctransrec","Creditors",1);        
        uhmd = getAnomoliesUninv("uninv","Uninv_Debtors",2);       
        uhmc = getAnomoliesUninv("cuninv","Uninv_Creditors",3);  
        
         List ls;
         List ls2,ls3,ls4;
         
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short)12);
        font.setFontName("Cambria");
        
        Font hfont = workbook.createFont();
        hfont.setFontHeightInPoints((short)12);
        hfont.setFontName("Cambria");
        
        CellStyle style = workbook.createCellStyle();
        CellStyle hstyle = workbook.createCellStyle();
        
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);  
        
        hstyle.setFont(font);
        hstyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        hstyle.setFillPattern(CellStyle.SOLID_FOREGROUND); 
        
         //System.out.println(hm);
         XSSFSheet sheet = workbook.createSheet("Debtors");
         XSSFSheet sheet2 = workbook.createSheet("Creditors");
         XSSFSheet sheet3 = workbook.createSheet("Uninv_Debtors");
         XSSFSheet sheet4 = workbook.createSheet("Uninv_Creditors");
         
         Iterator iterator = hmd.keySet().iterator();
            
         
                    int rowCount = 0;
                    Row row = sheet.createRow(0);
                    Cell hcell = row.createCell(1);
                    hcell.setCellValue((String) "RPPS");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(2);
                    hcell.setCellValue((String) "SVC");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(3);
                    hcell.setCellValue((String) "OPEN");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(4);
                    hcell.setCellValue((String) "RINVOICE");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(5);
                    hcell.setCellValue((String) "CORRECTION");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(6);
                    hcell.setCellValue((String) "ADJUST");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(7);
                    hcell.setCellValue((String) "O1C");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(8);
                    hcell.setCellValue((String) "SETTLED");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(9);
                    hcell.setCellValue((String) "ALLOC");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(10);
                    hcell.setCellValue((String) "WRITEOFF");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(11);
                    hcell.setCellValue((String) "CLOSE");
                    hcell.setCellStyle(hstyle);
                    hcell = row.createCell(12);
                    hcell.setCellValue((String) "RECDIFF");
                    hcell.setCellStyle(hstyle);
                    
                    
                    
                    while (iterator.hasNext()) {
                        
                        row = sheet.createRow(++rowCount);
                        int columnCount = 0;
                        String field = "";
                       
                        
                       String key = iterator.next().toString();
                       field = key;
                       Cell cell = row.createCell(++columnCount);
                      
                       if (field instanceof String) {
                                cell.setCellValue((String) field);
                                cell.setCellStyle(style);
                                     }
                       
                       //System.out.println(hm.get(key));
                       ls = (List<String>) hmd.get(key);    
                       System.out.println(ls);
                        Iterator<String> ite = ls.iterator();
                         while(ite.hasNext()){
                             field = ite.next();
                        //System.out.println( ite.next() );
                         cell = row.createCell(++columnCount);
                         if (field instanceof String) {
                                cell.setCellValue((String) field);
                                     } 
                         
                         cell.setCellStyle(style);
                         
                        }
                         
                          
                       
                    }
         
         
       Iterator iterator2 = hmc.keySet().iterator(); 
       rowCount = 0;
       
       while (iterator2.hasNext()) {
           
            row = sheet2.createRow(++rowCount);
                        int columnCount = 0;
                        String field = "";
                       
                        
                       String key = iterator2.next().toString();
                       field = key;
                       Cell cell = row.createCell(++columnCount);
                       if (field instanceof String) {
                                cell.setCellValue((String) field);
                                cell.setCellStyle(style);
                                     }
                       
                       //System.out.println(hm.get(key));
                       ls2 = (List<String>) hmc.get(key);    
                       System.out.println(ls2);
                        Iterator<String> ite = ls2.iterator();
                         while(ite.hasNext()){
                             field = ite.next();
                        //System.out.println( ite.next() );
                         cell = row.createCell(++columnCount);
                         if (field instanceof String) {
                                cell.setCellValue((String) field);
                                cell.setCellStyle(style);
                                     } 
                         
                        }
           
       }
       
       // uninv debtors
       
       Iterator iterator3 = uhmd.keySet().iterator(); 
       rowCount = 0;
       
       while (iterator3.hasNext()) {
           
            row = sheet3.createRow(++rowCount);
                        int columnCount = 0;
                        String field = "";
                       
                        
                       String key = iterator3.next().toString();
                       field = key;
                       Cell cell = row.createCell(++columnCount);
                       if (field instanceof String) {
                                cell.setCellValue((String) field);
                                cell.setCellStyle(style);
                                     }
                       
                       //System.out.println(hm.get(key));
                       ls3 = (List<String>) hmc.get(key);    
                       System.out.println(ls3);
                        Iterator<String> ite = ls3.iterator();
                         while(ite.hasNext()){
                             field = ite.next();
                        //System.out.println( ite.next() );
                         cell = row.createCell(++columnCount);
                         if (field instanceof String) {
                                cell.setCellValue((String) field);
                                cell.setCellStyle(style);
                                     } 
                         
                        }
           
       }
       
       
       // uninv creditors
       
       
       
       Iterator iterator4 = uhmc.keySet().iterator(); 
       rowCount = 0;
       
       while (iterator4.hasNext()) {
           
            row = sheet4.createRow(++rowCount);
                        int columnCount = 0;
                        String field = "";
                       
                        
                       String key = iterator4.next().toString();
                       field = key;
                       Cell cell = row.createCell(++columnCount);
                       if (field instanceof String) {
                                cell.setCellValue((String) field);
                                cell.setCellStyle(style);
                                     }
                       
                       //System.out.println(hm.get(key));
                       ls4 = (List<String>) uhmc.get(key);    
                       System.out.println(ls4);
                        Iterator<String> ite = ls4.iterator();
                         while(ite.hasNext()){
                             field = ite.next();
                        //System.out.println( ite.next() );
                         cell = row.createCell(++columnCount);
                         if (field instanceof String) {
                                cell.setCellValue((String) field);
                                cell.setCellStyle(style);
                                     } 
                         
                        }
           
       }
        /*
            Object[][] bookData = {
                {"Head First Java", "Kathy Serria", 79},
                {"Effective Java", "Joshua Bloch", 36},
                {"Clean Code", "Robert martin", 42},
                {"Thinking in Java", "Bruce Eckel", 35},
        };

        
        
        
        
         //int rowCount = 0;
        for (Object[] aBook : bookData) {
            Row row = sheet.createRow(++rowCount);
             
            int columnCount = 0;
             
            for (Object field : aBook) {
                Cell cell = row.createCell(++columnCount);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
             
        }

        */
       // workbook.setSheetName(n, WorkbookUtil.createSafeSheetName("Debtors"));
        try (FileOutputStream outputStream = new FileOutputStream("anomolies/GBRCN_Anomolies_27April.xlsx")) {
            workbook.write(outputStream);
        }
        
        
        
    }
     
     
     
     public static  HashMap<String, List<String>> getAnomoliesUninv(String tab,String type,Integer n) throws IOException {
        
        Connection con;
        Statement stmt;
        String sql;
        String rpps;  
        String svc;
        String open;
        String rinvoice;
        String correction;
        String adjust;
        String o1cf;
        String cdata;
        String recdiff;
        
        HashMap<String, List<String>> hashMap = new HashMap<String, List<String>>();
        
        try {
            
        con = DBDConnection.getConnection(tab);
        stmt = con.createStatement();
        
        
        
         sql = "select * from anomoly where recdiff<> 0 order by svc,recdiff";
                ResultSet rs = stmt.executeQuery(sql);
                rs = stmt.executeQuery(sql);
                
                
                while (rs.next())
                {
                 
                 rpps = rs.getString("rpps"); 
                 svc = rs.getString("svc");
                 open = rs.getString("open");
                 rinvoice = rs.getString("rinvoice");
                 correction = rs.getString("correction");
                 adjust = rs.getString("adjust");
                 o1cf = rs.getString("o1cf");
                 cdata = rs.getString("cdata");
                 recdiff = rs.getString("recdiff");
                 
                 hashMap.put(rpps,Arrays.asList(svc,open,cdata,rinvoice,correction,adjust,o1cf,recdiff));
                 
                               
                 
                 
                }
                
                rs.close();
                //System.out.println(hashMap);
                
               
                
                
                
                
        
        }
        catch (SQLException se){
            se.printStackTrace();
        }
        
        
        return hashMap;
        
    }
     
     
     public static  HashMap<String, List<String>> getAnomolies(String tab,String type,Integer n) throws IOException {
        
        Connection con;
        Statement stmt;
        String sql;
        String rpps;  
        String svc;
        String open;
        String rinvoice;
        String correction;
        String adjust;
        String o1cf;
        String settled;
        String alloc;
        String writeoff;
        String recdiff;
        String close;
        
        HashMap<String, List<String>> hashMap = new HashMap<String, List<String>>();
        
        try {
            
        con = DBDConnection.getConnection(tab);
        stmt = con.createStatement();
        
        
        
         sql = "select * from anomoly where recdiff<> 0 order by svc,recdiff";
                ResultSet rs = stmt.executeQuery(sql);
                rs = stmt.executeQuery(sql);
                
                
                while (rs.next())
                {
                 
                 rpps = rs.getString("rpps"); 
                 svc = rs.getString("svc");
                 open = rs.getString("open");
                 rinvoice = rs.getString("rinvoice");
                 correction = rs.getString("correction");
                 adjust = rs.getString("adjust");
                 o1cf = rs.getString("o1cf");
                 settled = rs.getString("settled");
                 alloc = rs.getString("alloc");
                 writeoff = rs.getString("writeoff");
                 close = rs.getString("close");
                 recdiff = rs.getString("recdiff");
                 
                 hashMap.put(rpps,Arrays.asList(svc,open,rinvoice,correction,adjust,o1cf,settled,alloc,writeoff,close,recdiff));
                 
                               
                 
                 
                }
                
                rs.close();
                //System.out.println(hashMap);
                
               
                
                
                
                
        
        }
        catch (SQLException se){
            se.printStackTrace();
        }
        
        
        return hashMap;
        
    }
    
}
