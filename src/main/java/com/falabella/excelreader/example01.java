/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.falabella.excelreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellType;

/**
 *
 * @author ext_ealinares
 */
public class example01 {

    //private row
    private int current_row;

    public static void main(String[] args) {
        // TODO code application logic here

        try {

            FileInputStream input_file = new FileInputStream(new File("C:\\ExcelFile\\input_file.xls"));

            String _file_separator = System.getProperty("file.separator");

            String lv_inputFileName = "template.xls";

            String lv_template_path_file_folder = "C:\\ExcelFile\\";

            String lv_template_path_file = lv_template_path_file_folder.concat(_file_separator).concat(lv_inputFileName);

            FileInputStream template_file = new FileInputStream(new File(lv_template_path_file));

            //Get the workbook_input_file instance for XLS file 
            HSSFWorkbook workbook_input_file = new HSSFWorkbook(input_file);

            HSSFWorkbook workbook_template_file = new HSSFWorkbook(template_file);

            //Get first sheet from the workbook_input_file
            HSSFSheet sheet_input_file = workbook_input_file.getSheetAt(0);

            //Get sheet  from workbook_template_file
            HSSFSheet sheet_template_file = workbook_template_file.getSheetAt(0);

            //Iterate through each row_input_files from first sheet_input_file
            Iterator<Row> row_input_fileIterator = sheet_input_file.iterator();

            EngineProcess(row_input_fileIterator, sheet_template_file);

            FileOutputStream out = PrintResults(lv_template_path_file_folder, _file_separator, lv_inputFileName, workbook_template_file);
           
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();

        }

    }

    private static FileOutputStream PrintResults(String lv_template_path_file_folder, String _file_separator, String lv_inputFileName, HSSFWorkbook workbook_template_file) throws IOException, FileNotFoundException {
        String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
        String _file_path = lv_template_path_file_folder.concat(_file_separator).concat(timeStamp.concat("_" + lv_inputFileName));
        FileOutputStream out
                = new FileOutputStream(new File(_file_path));
        workbook_template_file.write(out);
        return out;
    }

    private static void EngineProcess(Iterator<Row> row_input_fileIterator, HSSFSheet sheet_template_file) {
        //  row_input_fileIterator.
        System.out.println("Inicio proceso principal..........");
        
        while (row_input_fileIterator.hasNext()) {
            
            Row row_input_file = row_input_fileIterator.next();
            
            if (skip_columns_title(row_input_file)) {
                
                HSSFRow row_template_file = sheet_template_file.createRow(row_input_file.getRowNum());
                //For each row_input_file, iterate through each columns
                Iterator<Cell> cell_input_fileIterator = row_input_file.cellIterator();
                
                while (cell_input_fileIterator.hasNext()) {
                    
                    Cell cell_input_file = cell_input_fileIterator.next();
                    
                    HSSFCell cell_template_file = row_template_file.createCell(cell_input_file.getColumnIndex());
                    
                    // cell_template_file.setCellType(CellType.STRING);
                    switch (cell_input_file.getCellType()) {
                        
                        case STRING:
                            
                            cell_template_file.setCellValue(cell_input_file.getStringCellValue());
                            break;
                            
                        case NUMERIC:
                            
                            cell_template_file.setCellValue(cell_input_file.getNumericCellValue());
                            break;
                            
                    }
                    
                }
            }
            
        }
        
        System.out.println("Inicio proceso principal....");
    }

    private static boolean skip_columns_title(Row row_input_file) {

        boolean isFirstRow = false;

        if (row_input_file.getRowNum() > 0) {

            isFirstRow = true;

        }

        return isFirstRow;
    }
}
