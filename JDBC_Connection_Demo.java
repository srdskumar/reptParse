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
/* JDBC_Connection_Demo.java */
 
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
 
public class JDBC_Connection_Demo
{
  /* static block is executed when a class is loaded into memory 
   * this block loads MySQL's JDBC driver
   */
  static
  {
    try
    {
      // loads com.mysql.jdbc.Driver into memory
      Class.forName("com.mysql.jdbc.Driver");
    } 
    catch (ClassNotFoundException cnf) 
    {
      System.out.println("Driver could not be loaded: " + cnf);
    }
  }
 
  public static void main(String[] args)
  {
    String connectionUrl = "jdbc:mysql://localhost:3306/uninv";
    String dbUser = "root";
    String dbPwd = "suresh76";
    Connection conn;
    ResultSet rs;
    String queryString = "SELECT ID, NAME FROM EXPTABLE";
    System.setProperty("jdbc.drivers", "com.mysql.jdbc.Driver");
    
 
    try
     {
      conn = DriverManager.getConnection(connectionUrl, dbUser, dbPwd);
      Statement stmt = conn.createStatement();
 
      // INSERT A RECORD
      stmt.executeUpdate("INSERT INTO EXPTABLE (NAME) VALUES (\"TINU K\")");
 
      // SELECT ALL RECORDS FROM EXPTABLE
      rs = stmt.executeQuery(queryString);
 
      System.out.println("ID \tNAME");
      System.out.println("============");
      while(rs.next())
      {
        System.out.print(rs.getInt("id") + ".\t" + rs.getString("name"));
        System.out.println();
      }
      if (conn != null)
      {
        conn.close();
        conn = null;
      }
    }
    catch (SQLException sqle) 
    {
      System.out.println("SQL Exception thrown: " + sqle);
    }
  }
} //JDBC_Connection_Demo ends here
 
