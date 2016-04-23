/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapp;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashMap;

/**
 *
 * @author g706134
 */
public class CompileAnomoliesDebtorTransrecEE {
    
     public static void main(String[] args)  {
         
         loadUniqueRpps();
         
     }
     
     public static void loadUniqueRpps()
     {
        Connection conn = null;
        Statement stmt = null;        
        Statement utmt = null;
        String sql = "";
        String rpps = "";
        double sdrval;
        String[] temp;
        String[] tab_list = new String[]{"open", "rinvoice","correction","adjust","o1cf","settled","alloc","writeoff","close"};
        
        try {
            
            Connection con = DBConnection.getConnection();
            stmt = con.createStatement();
            utmt = con.createStatement();
            HashMap<String, Integer> hmap = new HashMap<String, Integer>();
            int i = 0;
            int j = 0;
            
            
            for (String t: tab_list) {  
                
                sql = "select rpps from "+t;
                ResultSet rs = stmt.executeQuery(sql);
                rs = stmt.executeQuery(sql);
                
                
                while (rs.next())
                {
                 i++;
                 rpps = rs.getString("rpps");
                 hmap.put(rpps,i);  
                 System.out.println(t+"-->"+i+"-->"+rpps);
                 
                }
                
                rs.close();
                
            }
            
             for (String key : hmap.keySet()) {
                 j++;
                 //System.out.println(j+"--->" + key + "" + hmap.get(key));  
                 temp = key.split("-");
                 System.out.println("-->"+temp[3]+"<--");
                 if(temp[3].length() > 5 )
                 {
                     continue;
                 }
                 sql = "insert into anomoly(rpps,svc) values(\""+key+"\",\""+temp[3]+"\")";
                 System.out.println(sql);                 
                 utmt.executeUpdate(sql);
                 //sql = "update anomoly set open=(select sdrval from  open where rpps=\""+key+"\") where rpps=\""+key+"\"";
                 System.out.println(sql);                 
                 //utmt.executeUpdate(sql);
                 
                }
             
             for (String t: tab_list) {  
                
                sql = "select rpps from "+t;
                ResultSet rs = stmt.executeQuery(sql);
                rs = stmt.executeQuery(sql);
                
                
                while (rs.next())
                {
                 
                 rpps = rs.getString("rpps");
                 sql = "update anomoly set "+t+"=(select sum(sdrval) from  "+t+" where rpps=\""+rpps+"\") where rpps=\""+rpps+"\"";
                 //hmap.put(rpps,i);  
                 //System.out.println(t+"-->"+rpps+"--->");
                 System.out.println(sql);
                 utmt.executeUpdate(sql);
                 
                }
                
                rs.close();
                
            }
               
                sql = "update anomoly set recdiff=open + rinvoice + correction + adjust + o1cf - settled - alloc - writeoff - close";
                System.out.println(sql);                 
                utmt.executeUpdate(sql);
             


        }
        
        catch (SQLException se){
            se.printStackTrace();
        }
        
        
     }
    
    
}
