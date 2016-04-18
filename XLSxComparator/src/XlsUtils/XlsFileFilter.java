/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package XlsUtils;

import java.io.File;
import javax.swing.filechooser.FileFilter;

/**
 *
 * @author jagam
 */
public class XlsFileFilter extends FileFilter{
    
    @Override
    public boolean accept(File f) {
        return  f.isDirectory() || f.getName().endsWith(".xls") || f.getName().endsWith(".xlsx");
    }

    @Override
    public String getDescription() {
        return "xls/xlsx";
    }
}
