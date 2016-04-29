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

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;


public class generateAnomolies {
    
    
    public static void main(String[] args) throws IOException {
        
        
        Connection con = null;
        Statement stmt = null;
        String sql = "";
        String rpps = "";  
        String svc = "";
        String open = "";
        String rinvoice = "";
        String correction = "";
        String adjust = "";
        String o1cf = "";
        String settled = "";
        String alloc = "";
        String writeoff = "";
        String recdiff = "";
        
        try {
            
        con = DBDConnection.getConnection("transrec");
        stmt = con.createStatement();
        
         sql = "select rpps from anomoly where recdiff<> 0 order by svc,recdiff";
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
                 recdiff = rs.getString("recdiff");
                 
                         
                 
                 System.out.println(rpps);
                 
                }
                
                rs.close();
        
        }
        catch (SQLException se){
            se.printStackTrace();
        }
        
        generateAnomolyReport("Debtor");
        //generateAnomolyReport("Creditor");
        //generateAnomolyReport("Uninv_Debtor");
        //generateAnomolyReport("Uninv_creditor");
        
    }
    
    
     public static void generateAnomolyReport(String tab) throws IOException
    {
        ArrayList<String> ar = new ArrayList<String>();
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Java Books");
         
        Object[][] bookData = {
                {"Head First Java", "Kathy Serria", 79},
                {"Effective Java", "Joshua Bloch", 36},
                {"Clean Code", "Robert martin", 42},
                {"Thinking in Java", "Bruce Eckel", 35},
        };
        
         int rowCount = 0;
         
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
        
        try (FileOutputStream outputStream = new FileOutputStream("anomolies/GBRCN_Anomolies_27April.xlsx")) {
            workbook.write(outputStream);
        }
        
        
        
    }
    
}
