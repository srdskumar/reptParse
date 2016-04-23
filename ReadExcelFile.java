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


public class ReadExcelFile {
    
   
public static void main(String[] args) {

        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream("GBRCNCOR.xlsx");

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            // for each sheet in the workbook
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                System.out.println("Sheet name: " + workbook.getSheetName(i));
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fileInputStream != null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    
}
