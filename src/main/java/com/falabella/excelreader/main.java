/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.falabella.excelreader;

import java.io.FileInputStream;
import java.util.Properties;

/**
 *
 * @author ext_ealinares
 */
public class main {

    /**
     * @param args the command line arguments
     * @throws java.lang.Exception
     */
    public static void main(String[] args) throws Exception {
        // TODO code application logic here

        if (args.length == 0) {

            throw new Exception("Parametro nombre archivo requerido! ");

        }

        FileManager fm = new FileManager();

        String lv_param = args[0];

        // lv_param = "cotas_maestro_base_20180828.xls";
        System.out.println("Inicio proceso archivo: " + lv_param);

        fm.main(lv_param);

    }

}
