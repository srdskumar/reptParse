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


import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.ResultSet;
import java.util.Iterator;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.lang.*;




public class SanityCheckCreditorTransRec {
    
    
    public static void main(String[] args)  {
        
        Connection conn = null;
        Statement stmt = null;
        
        String sql = "";
        
         String[] col_list = new String[]{"open", "rinvoice","correction","adjust","o1cf","close","settled","writeoff","alloc"};
         String[] svclist = new String[]{"Voice", "SMS", "IOT","ALL"};
                  
        
        try {
            
            Connection con = DBConnection.getConnection();
            stmt = con.createStatement();
            
            double total;
            HashMap<String, Double> hmap = new HashMap<String, Double>();
            
            
            
            
            for (String s: svclist) {           
                    
                    System.out.println("----------------->"+s+"<---------------");
                    
                    for (String t: col_list) {           
                    //Do your stuff here
                    sql = "select sum(sdrval) total from "+t;
                    if(!s.equals("ALL")) 
                    { 
                        sql = "select sum(sdrval) total from "+t+" where rpps like '%"+s+"%'";
                        
                    }
                    
                    ResultSet rs = stmt.executeQuery(sql);
                    rs = stmt.executeQuery(sql);
                    rs.next();  
                    total = rs.getDouble("total");
                    //System.out.print(sql); 
                    //System.out.print("-->"); 
                    //System.out.print(total); 
                    //System.out.print("-->");
                    hmap.put(t,total);   
                    }
                    
                    sql = "insert into sanity (svc,open,rinvoice,correction,adjust,o1cf,settled,alloc,writeoff,close) values(\""+s+"\","+hmap.get("open")+","+hmap.get("rinvoice")+","+hmap.get("correction")+","+hmap.get("adjust")+","+hmap.get("o1cf")+","+hmap.get("settled")+","+hmap.get("alloc")+","+hmap.get("writeoff")+","+hmap.get("close")+")";
                    System.out.println(sql);
                    stmt.execute(sql);
                    
                    
                   
                    
                    
                }
            
            
            
            
            
            
            
      stmt.close();
        }
        
        catch (SQLException se){
            se.printStackTrace();
        }
        
        findDiff();
        
        
    }
    
    
    public static void findDiff()
    {
        Connection conn = null;
        Statement stmt = null;
        Statement utmt = null;
        String sql = "";
        
        
        try {
            
             double open = 0;
             double cdata = 0;
             double rinvoice = 0;
             double correction = 0;
             double adjust = 0;
             double o1cf = 0;
             double settled = 0;
             double alloc = 0;
             double writeoff = 0;
             double close = 0;
             int id;
             double recdiff;
             String svc;
            
             Connection con = DBConnection.getConnection();
             stmt = con.createStatement();  
             utmt = con.createStatement();
             
             sql = "select * from sanity";
             ResultSet rs = stmt.executeQuery(sql);
             rs = stmt.executeQuery(sql);
             
             while (rs.next())
             {
             //rs.next(); 
             id = rs.getInt("id");
             svc = rs.getString("svc");
             //Additions
             open = rs.getDouble("open");
             //cdata = rs.getDouble("cdata");
             correction = rs.getDouble("correction");
             adjust = rs.getDouble("adjust");
             o1cf = rs.getDouble("o1cf");
             settled = rs.getDouble("settled");
             alloc = rs.getDouble("alloc");
             writeoff = rs.getDouble("writeoff");
             
             //cdata + correction + adjust + o1cf + 
             
             //Netoffs
             rinvoice = rs.getDouble("rinvoice");
             close = rs.getDouble("close");
             
             recdiff = open - rinvoice + correction + adjust + o1cf - settled + alloc + writeoff - close;
             recdiff = Math.round(recdiff);
             
             
             
             sql = "update sanity set recdiff = "+recdiff+" where id="+id;             
             utmt.executeUpdate(sql);
             System.out.println(sql);
             }
             
        }
        catch (SQLException se){
            se.printStackTrace();
        }
        
        
    }
    
    
}
