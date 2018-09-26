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
public class ReadExcelFile {

    public static void main(String[] args) {
        // TODO code application logic here

        try {

            FileInputStream file01 = new FileInputStream(new File("C:\\ExcelFile\\input_file.xls"));

            String _file_separator = System.getProperty("file.separator");

            String lv_inputFileName = "template.xls";

            String lv_path_folder = "C:\\ExcelFile\\";

            String lv_path = lv_path_folder.concat(_file_separator).concat(lv_inputFileName);

            FileInputStream file02 = new FileInputStream(new File(lv_path));

            //Get the workbook01 instance for XLS file 
            HSSFWorkbook workbook01 = new HSSFWorkbook(file01);

            HSSFWorkbook workbook02 = new HSSFWorkbook(file02);

            //Get first sheet01 from the workbook01
            HSSFSheet sheet01 = workbook01.getSheetAt(0);

            //set sheet02 from workbook 2
            HSSFSheet sheet02 = workbook02.getSheetAt(0);

            HSSFRow row01 = sheet02.getRow(3);
            if (row01 == null) {
                row01 = sheet02.createRow(2);
            }
            HSSFCell cell02 = row01.getCell(3);
            if (cell02 == null) {
                cell02 = row01.createCell(3);
            }
            cell02.setCellType(CellType.STRING);

            //Iterate through each rows from first sheet01
            Iterator<Row> rowIterator = sheet01.iterator();

            String celda = null;
            int fila = 1;
            int col = 0;

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                row01 = sheet02.createRow(fila);
                //For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t\t");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t\t");
                            cell02 = row01.createCell(3);
                            cell02.setCellValue(cell.getStringCellValue());
                            break;

                    }

                    col++;
                }

                System.out.println("");
                fila++;
            }

            String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());

            String _file_path = lv_path_folder.concat(_file_separator).concat(timeStamp.concat("_" + lv_inputFileName));

            FileOutputStream out
                    = new FileOutputStream(new File(_file_path));
            workbook02.write(out);
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();

        }
    }

}
