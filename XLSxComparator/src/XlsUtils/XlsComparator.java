/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package XlsUtils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jagam
 */
public class XlsComparator {
    
    private XlsComparator(){}
    
    public static boolean comparaExcel(Workbook excel1, Workbook excel2, StringBuilder cache){
        boolean res = true;
        int numSheet1 = excel1.getNumberOfSheets();
        int numSheet2 = excel2.getNumberOfSheets();
        int maxNumSheets = numSheet1 > numSheet2?numSheet1:numSheet2;
        
        try{
            for( int i = 0 ; i < maxNumSheets ; i++ )
                if (! comparaHoja(excel1.getSheetAt(i), excel2.getSheetAt(i), cache) )
                    res = false;
        }catch( IllegalArgumentException|NullPointerException e ){
            res = false;
        }
        
        return res;
    }
    
    public static boolean comparaHoja(Sheet hoja1, Sheet hoja2, StringBuilder cache){
        boolean res = true;
        int numRows1 = hoja1.getPhysicalNumberOfRows();
        int numRows2 = hoja2.getPhysicalNumberOfRows();
        int maxNumRows = numRows1 > numRows2?numRows1:numRows2;
        
        try{
            for( int i = 0 ; i < maxNumRows ; i++ ){
                if( ! comparaFila( hoja1.getRow(i), hoja2.getRow(i), cache ) )
                    res = false;
            }
        }catch(IllegalArgumentException|NullPointerException e){
            res = false;
        }
        
        return res;
    }
    
    public static boolean comparaFila(Row fila1, Row fila2, StringBuilder cache){
        boolean res = true;
        
        int numCell1 = fila1.getPhysicalNumberOfCells();
        int numCell2 = fila2.getPhysicalNumberOfCells();
        int maxNumCells = numCell1 > numCell2?numCell1:numCell2;
        
        try{
            for( int i = 0 ; i < maxNumCells ; i++ ){
                if( ! comparaCelda( fila1.getCell(i), fila2.getCell(i), cache ) )
                    res = false;
            }
        }catch(IllegalArgumentException|NullPointerException e){
            res = false;
        }
        
        return res;
    }
    
    public static boolean comparaCelda(Cell celda1, Cell celda2, StringBuilder cache){
        Object value1 = getCellValue(celda1);
        Object value2 = getCellValue(celda2);
        String adress;
        boolean res;
        
        try{
            res = value1.equals(value2);
            adress = celda1.getAddress().formatAsString();
        }catch(NullPointerException ex){
            res = value2.equals(value1);
            adress = celda2.getAddress().formatAsString();
        }
        
        if( !res )
            cache.append("El valor de ").append(adress).append(" es diferente en los 2 excel:\n")
                    .append("Excel 1:").append(value1)
                    .append("\n")
                    .append("Excel 2:").append(value2)
                    .append("\n");
            
        return res;
    }
    
    
    /**
     * Obtiene el valor de una Cell de Excel
     * @param cell
     * @return 
     */
    public static Object getCellValue(Cell cell) {
        Object result = null;
        
        if(cell!=null){
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        result = cell.getDateCellValue();
                    } else {
                        if(cell.getNumericCellValue() == (int) cell.getNumericCellValue()){
                            result = (int)cell.getNumericCellValue();
                        }else{
                            result = cell.getNumericCellValue();
                        }
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    result = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    result = cell.getBooleanCellValue();
                    break;
                case Cell.CELL_TYPE_FORMULA:
                case Cell.CELL_TYPE_BLANK:
                case Cell.CELL_TYPE_ERROR:
                    result = null;
                    break;
            }
        }

        return result;
    }
}
