/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package main;

import XlsUtils.XlsFileFilter;
import XlsUtils.XlsComparator;
import XlsUtils.XlsFactory;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Scanner;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jagam
 */
public class Main {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        Scanner scan = new Scanner(System.in);
        StringBuilder cache = new StringBuilder();
        File first;
        File second;
        File logPath;
        try{

            JOptionPane.showMessageDialog(null, "Dame la ruta de un fichero excel");
            first = getAnyFile();
            JOptionPane.showMessageDialog(null, "Dame la ruta de otro fichero excel");
            second = getAnyFile();
            JOptionPane.showMessageDialog(null, "Dame la ruta donde guardar el log");
            logPath = getAnyFile();

            try( Workbook excel1 = XlsFactory.getWorkbook(first); Workbook excel2 = XlsFactory.getWorkbook(second) ){

                XlsComparator.comparaExcel(excel1, excel2, cache);

                System.out.println( cache.toString().isEmpty()?"Todo es correcto":cache );
            }

            try( FileWriter fw = new FileWriter(logPath) ){
                String res = cache.toString().replace("$Excel1$", first.getName()).replace("$Excel2$", second.getName());
                fw.append( res );
            }
            
        }catch(NullPointerException|IOException ex){
            JOptionPane.showMessageDialog(null, ex.getStackTrace(), ex.getLocalizedMessage(), JOptionPane.ERROR_MESSAGE);
        }
    }
    
    public static File getAnyFile(){
        File res;
        JFileChooser fileChooser = new JFileChooser();
        
        fileChooser.setFileFilter( new XlsFileFilter() );
        
        int i = fileChooser.showOpenDialog(null);
        
        if( i == JFileChooser.APPROVE_OPTION )
            res = fileChooser.getSelectedFile();
        else
            throw new NullPointerException("No se ha seleccionado ningún archivo");
        
        return res;
    }
    
    
}
