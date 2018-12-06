/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.falabella.services;

import com.falabella.excelreader.Main;
import static com.falabella.services.CopySheets.copySheets;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author ext_ealinares
 */
public class FileManager extends CopySheets {

    private String template_file_path;
    private String template_file_name;
    private String input_file_folder;
    private String output_file_folder;
    private Properties props;
    private String _file_path;
    private String _extension_file;

    public FileManager() {

        super();
        this.props = new Properties();
        String _prop_path = "src/resources/properties/properties.properties";

        try {
            props.load(new FileInputStream(_prop_path));
            this.setTemplate_file_path(props.getProperty("template_path_file").trim());
            this.setTemplate_file_name(props.getProperty("template_file_name").trim());
            this.setInput_file_folder(props.getProperty("input_file_folder").trim());
            this.setOutput_file_folder(props.getProperty("output_file_folder").trim());

        }
        catch (Exception e) {
            System.out.println("No se pudo rescatar archivo de propiedades.\n" + _prop_path);
            e.printStackTrace();
        }
    }

    public String getTemplate_file_path() {

        return template_file_path;
    }

    public void setTemplate_file_path(String template_file_path) {
        this.template_file_path = template_file_path;
    }

    public String getInput_file_folder() {
        return input_file_folder;
    }

    public void setInput_file_folder(String input_file_folder) {
        this.input_file_folder = input_file_folder;
    }

    public String getOutput_file_folder() {
        return output_file_folder;
    }

    public void setOutput_file_folder(String output_file_folder) {
        this.output_file_folder = output_file_folder;
    }

    public String getTemplate_file_name() {
        return template_file_name;
    }

    public void setTemplate_file_name(String template_file_name) {
        this.template_file_name = template_file_name;
    }

    public String getFile_path() {
        return _file_path;
    }

    public void setFile_path(String _prop_name, String _input_file) {

        String _file_separator = null;

        String _os_name = System.getProperty("os.name").toUpperCase();

        if (_os_name.contains("WINDOWS")) {

            _file_separator = "/";
        }
        else {

            _file_separator = System.getProperty("file.separator");
        }

        //String _file_path = _prop_name.concat(_file_separator).concat(_input_file);
        this._file_path = _prop_name.concat(_file_separator).concat(_input_file);

        // this._file_path = _file_path;
    }

    public String getExtension_file() {
        return _extension_file;
    }

    public void setExtension_file(String _extension_file) {
        this._extension_file = _extension_file;
    }

    public void CopyDataBetweenSheets(File myFile) {

        try {
            // FileInputStream template_file = new FileInputStream(new File(""));

            Workbook input_Workbook = WorkbookFactory.create(myFile);

            this.setFile_path(this.getTemplate_file_path(), this.getTemplate_file_name());

            // FileInputStream template_fis = new FileInputStream(new File(this.getFile_path()));
            File templateFile = new File(this.getFile_path());

            Workbook template_Workbook = WorkbookFactory.create(templateFile);

            //Sheet input_sheet = input_Workbook.getSheetAt("TRANSIT_TIME_LANE");
            //  Sheet input_sheet = input_Workbook.getSheet("TRANSIT_TIME_LANE");
            //    Sheet template_sheet = template_Workbook.getSheet("TRANSIT_TIME_LANE");
            //InputStream targetStream = FileUtils.openInputStream(myFile);
            // List<InputStream> myList = new List<InputStream>;
            //myList
            // FileInputStream fis = new FileInputStream(myFile);
            InputStream file = new FileInputStream(myFile);

            List<InputStream> myList = new ArrayList<InputStream>();

            myList.add(file);

            Workbook output_Workbook = mergeExcelFiles(template_Workbook, myList);

            //  copySheets(template_sheet, input_sheet);
            setFile_path(this.getOutput_file_folder(), myFile.getName());

            FileOutputStream out
                    = new FileOutputStream(new File(this.getFile_path()));
            output_Workbook.write(out);

        }
        catch (FileNotFoundException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }
        catch (IOException ex) {
            Logger.getLogger(FileManager.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    public void CopyDataBetweenSheets2(File myFile) {

        try {
            // FileInputStream template_file = new FileInputStream(new File(""));

            Workbook input_Workbook = WorkbookFactory.create(myFile);

            this.setFile_path(this.getTemplate_file_path(), this.getTemplate_file_name());

            // FileInputStream template_fis = new FileInputStream(new File(this.getFile_path()));
            File templateFile = new File(this.getFile_path());

            Workbook template_Workbook = WorkbookFactory.create(templateFile);

            Sheet input_sheet = input_Workbook.getSheet("TRANSIT_TIME_LANE");

            Sheet template_sheet = template_Workbook.getSheet("TRANSIT_TIME_LANE");

            MainProcess(input_sheet, template_sheet);

            input_sheet = input_Workbook.getSheet("TRANSIT_LANE_DTL");

            template_sheet = template_Workbook.getSheet("TRANSIT_LANE_DTL");
            
            MainProcess(input_sheet, template_sheet);
              
            setFile_path(this.getOutput_file_folder(), myFile.getName());

            FileOutputStream out
                    = new FileOutputStream(new File(this.getFile_path()));
            template_Workbook.write(out);

        }
        catch (FileNotFoundException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }
        catch (IOException ex) {
            Logger.getLogger(FileManager.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    private void MainProcess(Sheet input_sheet, Sheet template_sheet) {
        for (Row row : input_sheet) {
            
            if (skip_columns_title(row)) {
                
                
                Row row_template_file = template_sheet.createRow(row.getRowNum());
                
                for (Cell cell : row) {
                    
                    
                    Cell cell_template_file = row_template_file.createCell(cell.getColumnIndex());
                    
                    switch (cell.getCellType()) {
                        
                        case STRING:
                            cell_template_file.setCellValue(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            cell_template_file.setCellValue(cell.getNumericCellValue());
                            break;
                            
                        default:
                            break;
                            
                    }
                    
                }
                
            }
            
        }
    }

    private boolean skip_columns_title(Row row_input_file) {

        boolean isFirstRow = false;

        if (row_input_file.getRowNum() > 1) {

            isFirstRow = true;

        }

        return isFirstRow;
    }

}
