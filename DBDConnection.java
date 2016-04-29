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

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;


public class DBDConnection {
    
public static Connection getConnection(String dbs) {
        Properties props = new Properties();
        FileInputStream fis = null;
        Connection con = null;
       
        try {
            fis = new FileInputStream("db.properties");
            props.load(fis);
             String dburl = props.getProperty("DBD_URL")+dbs;
        System.out.println("---------------->");
        System.out.println(dburl);
            // load the Driver Class
            Class.forName(props.getProperty("DB_DRIVER_CLASS"));
 
            // create the connection now
            con = DriverManager.getConnection(dburl,
                    props.getProperty("DB_USERNAME"),
                    props.getProperty("DB_PASSWORD"));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return con;
    } 
    
    
}
