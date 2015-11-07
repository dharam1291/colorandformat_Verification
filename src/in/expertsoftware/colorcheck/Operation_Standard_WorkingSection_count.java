/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package in.expertsoftware.colorcheck;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dharam
 */
public class Operation_Standard_WorkingSection_count {
    public ArrayList<String>working_Section(int SOOWDLocation,int EOOWDLocation,XSSFWorkbook workbook)
    {
     ArrayList<String> workingSection=new ArrayList<String>();     
     XSSFRow row;
     XSSFCell cell;
     XSSFSheet Sheet=workbook.getSheet("Operation_Standard");
     String genrateFormula;
     for(int start=(SOOWDLocation-1);start<EOOWDLocation;start++)
        {
            try
                {
                row = Sheet.getRow(start);
                cell = row.getCell(2);
                switch (cell.getCellType()) 
                        {                        
                            case Cell.CELL_TYPE_STRING:
                                genrateFormula="Operation_Standard!C"+(row.getRowNum()+1);
                                workingSection.add(genrateFormula);
                                break;
                             case Cell.CELL_TYPE_BLANK:
                                break;
                              default:
                                System.out.println("Error");
                                break;
                        }
                }
            catch(Exception e){
                System.out.println(e.getMessage());
                e.printStackTrace();
            }
       }
     return workingSection;    
    }
    
}
