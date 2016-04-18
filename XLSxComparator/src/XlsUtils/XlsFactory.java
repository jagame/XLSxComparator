/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package XlsUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author jagam
 */
public class XlsFactory {
    
    private XlsFactory(){
        
    }
    
    public static Workbook getWorkbook(File file) throws FileNotFoundException, IOException{
        Workbook res;
        if( file.isFile() ){
            if( file.getName().endsWith(".xls") )
                res = new HSSFWorkbook( new FileInputStream(file) );
            else if( file.getName().endsWith(".xlsx") )
                res = new XSSFWorkbook( new FileInputStream(file) );
            else
                throw new FileNotFoundException("No se reconoce el fichero como un archivo excel");
        }else
            throw new FileNotFoundException("No se reconoce la ruta indicada como un fichero");
                
        return res;
    }
    
}
