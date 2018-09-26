/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.falabella.excelreader;

import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;

/**
 *
 * @author ext_ealinares
 */
public class FileManager {

    private String template_file_path;
    private String template_file_name;
    private String input_file_folder;
    private String output_file_folder;
    private String _file_path;

    public void main(String _file_name) {

        this.getconfig();

        set_file_path(this.template_file_path, getTemplate_file_name());
        String _template_file = get_file_path();
        set_file_path(this.input_file_folder, _file_name);
        String _input_file = get_file_path();

        if (check_file_exists(_template_file)
                && check_file_exists(_input_file)) {

            try {

                ReadWriteWorkbook.main(_template_file, _input_file, output_file_folder);
                System.out.println("fin proceso excel");

            } catch (Exception ioe) {
                ioe.printStackTrace();
                System.out.println(ioe);
            }

        } else {

            System.out.println("Verifique archivos: ");
            System.out.println("Plantilla: " + _template_file);
            System.out.println("input: " + _input_file);

        }

    }

    public String get_file_path() {
        return _file_path;
    }

    public void set_file_path(String _prop_name, String _input_file) {

        String _file_separator = System.getProperty("file.separator");
        String _file_path = _prop_name.concat(_file_separator).concat(_input_file);

        this._file_path = _file_path;
    }

    private boolean check_file_exists(String filePathString) {

        File f = new File(filePathString);
        boolean isfile = false;

        try {
            if (f.isFile() && f.exists()) {

                isfile = true;
            }
        } catch (Exception e) {
            System.out.println("No se pudo rescatar archivo solicitado.\n");
            e.printStackTrace();
        }

        return isfile;

    }

    private void getconfig() {

        //String _prop_path = "src/properties/resources.properties";
        String _prop_path = "resources/properties/properties.properties";

        Properties props = new Properties();

        try {
            props.load(new FileInputStream(_prop_path));
            this.template_file_path = props.getProperty("template_path_file").trim();
            this.setTemplate_file_name(props.getProperty("template_file_name").trim());
            this.input_file_folder = props.getProperty("input_file_folder").trim();
            this.output_file_folder = props.getProperty("output_file_folder").trim();

        } catch (Exception e) {
            System.out.println("No se pudo rescatar archivo de propiedades.\n" + _prop_path);
            e.printStackTrace();
        }

    }

    public String getTemplate_file_name() {
        return template_file_name;
    }

    public void setTemplate_file_name(String template_file_name) {
        this.template_file_name = template_file_name;
    }

}
