/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.falabella.excelreader;

import com.falabella.services.FileManager;
import java.io.File;
import java.io.PrintStream;
import org.apache.commons.io.FilenameUtils;


/**
 *
 * @author ext_ealinares
 */
public class Main {
    
    private static FileManager fmt;
  
  public static void main(String[] args)
    throws Exception
  {
    fmt = new FileManager();
    
    File folder = new File(fmt.getInput_file_folder());
    listFilesForFolder(folder);
  }
  
  public static void listFilesForFolder(File folder)
  {
    for (File fileEntry : folder.listFiles()) {
      if (fileEntry.isDirectory()) {
        listFilesForFolder(fileEntry);
      } else {
        do_file_process(fileEntry);
      }
    }
  }
  
  public static void do_file_process(File fileEntry)
  {
    String ext1 = FilenameUtils.getExtension(fileEntry.getAbsolutePath());
    if (ext1.equals("csv")) {
      move_file(fileEntry);
    } else if (ext1.equals("xls")) {
      plantilla(fileEntry);
    }
  }
  
  private static void move_file(File myFile)
  {
    System.out.println("Moviendo csv");
    System.out.println(fmt.getOutput_file_folder());
    fmt.setFile_path(fmt.getOutput_file_folder(), myFile.getName());
    myFile.renameTo(new File(fmt.getFile_path()));
  }
  
  private static void plantilla(File myFile)
  {
    System.out.println("Procesando plantilla xls");
    
    fmt.CopyDataBetweenSheets2(myFile);
  }
}
