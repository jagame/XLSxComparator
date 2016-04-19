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
        int numRows1 = hoja1.getLastRowNum(); // Te devuelve el índice de la última fila
        int numRows2 = hoja2.getLastRowNum();
        int maxNumRows = numRows1 > numRows2?numRows1:numRows2;
        Row row1;
        Row row2;
        
        for( int i = 0 ; i <= maxNumRows ; i++ ){
            try{
                row1 = hoja1.getRow(i);
            }catch(NullPointerException|IllegalArgumentException ex){
                row1 = null;
            }
            try{
                row2 = hoja2.getRow(i);
            }catch(NullPointerException|IllegalArgumentException ex){
                row2 = null;
            }
            if( ! comparaFila( row1, row2, cache ) )
                res = false;
        }
        
        return res;
    }
    
    public static boolean comparaFila(Row fila1, Row fila2, StringBuilder cache){
        boolean res = true;
        int numCell1;
        int numCell2;
        Cell cell1;
        Cell cell2;
        
        try{
            numCell1 = fila1.getLastCellNum(); // Te devuelve el índice de la última celda MÁS 1
        }catch( NullPointerException e){
            numCell1 = 0;
        }
        try{
            numCell2 = fila2.getLastCellNum();
        }catch(NullPointerException e){
            numCell2 = 0;
        }
        
        int maxNumCells = numCell1 > numCell2?numCell1:numCell2;
        
        for( int i = 0 ; i < maxNumCells ; i++ ){
            try{
                cell1 = fila1.getCell(i);
            }catch(NullPointerException|IllegalArgumentException ex){
                cell1 = null;
            }
            try{
                cell2 = fila2.getCell(i);
            }catch(NullPointerException|IllegalArgumentException ex){
                cell2 = null;
            }
            if( ! comparaCelda( cell1, cell2, cache ) )
                res = false;
        }
        
        return res;
    }
    
    public static boolean comparaCelda(Cell celda1, Cell celda2, StringBuilder cache){
        Object value1 = getCellValue(celda1);
        Object value2 = getCellValue(celda2);
        String adress;
        boolean res;
        
//        Esta primera comparación nos libra de 3 casos, 1 de ellos problemático:
//        1) Son primitivos iguales por lo que no hay que hacer más gestión
//        2) Son el mismo objeto por lo que no hay que hacer más gestión
//        3) Son los 2 nulos, lo cual controlar podría ensuciar el código y realmente eso significa que son iguales y no hay que hacer más gestión
        if( value1 == value2 )
            res = true;
        else{
            try{
                res = value1.equals(value2);
                adress = celda1.getAddress().formatAsString();
            }catch(NullPointerException ex){
                res = value2.equals(value1);
                adress = celda2.getAddress().formatAsString();
            }

            if( cache != null && !res )
                cache.append("DEBUG:: El valor de ").append(adress).append(" es diferente en los 2 excel: ").append("$Excel1$: ").append(value1).append(" || ").append("$Excel2$: ").append(value2).append("\r\n");
        }
        
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
