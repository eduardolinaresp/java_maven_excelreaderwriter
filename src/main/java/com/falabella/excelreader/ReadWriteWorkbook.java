/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.falabella.excelreader;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;

/**
 *
 * @author ext_ealinares
 */
public class ReadWriteWorkbook {

    public static void main(String p_template_file, String p_input_file, String p_output_file ) throws IOException {
        FileInputStream fileIn = null;
        FileOutputStream fileOut = null;

        String _input_file = p_template_file;
        //String _input_file = p_template_file;

        try {
            // fileIn = new FileInputStream("workbook.xls");
            fileIn = new FileInputStream(_input_file);
            POIFSFileSystem fs = new POIFSFileSystem(fileIn);
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            
            HSSFSheet sheet = wb.getSheetAt(4);

            //HSSFRow row = sheet.getRow(2);
            HSSFRow row = sheet.getRow(3);
            if (row == null) {
                row = sheet.createRow(2);
            }
            HSSFCell cell = row.getCell(3);
            if (cell == null) {
                cell = row.createCell(3);
            }
            cell.setCellType(CellType.STRING);
            cell.setCellValue("a test");
            
            String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
      
            // Write the output to a file
            //fileOut = new FileOutputStream(timeStamp.concat("_cotas_maestro_base.xls"));
             String _file_separator = System.getProperty("file.separator");
             String _file_path = p_output_file.concat(_file_separator).concat(timeStamp.concat("_cotas_maestro_base.xls"));
        
            fileOut = new FileOutputStream(_file_path);
            
            wb.write(fileOut);

        } finally {
            if (fileOut != null) {
                fileOut.close();
            }
            if (fileIn != null) {
                fileIn.close();
            }
        }
    }

}
