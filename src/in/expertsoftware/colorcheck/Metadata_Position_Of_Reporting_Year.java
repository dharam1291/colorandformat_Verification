/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package in.expertsoftware.colorcheck;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 *
 * @author Dharam
 */
public class Metadata_Position_Of_Reporting_Year {
    public boolean verify_SORY_metadata_posiition( XSSFSheet Sheet,ArrayList<TokenModel> tokenmodel)
    {    
    int SORY_rowindex = -1;
    boolean correct=true;
    for(int i=0;i<tokenmodel.size();i++)
      {
         if((tokenmodel.get(i).token_name.equals("SORY")))
                 SORY_rowindex=tokenmodel.get(i).row_no;                 
      }
    if(SORY_rowindex==-1)
    {correct=false;}
    else
    {
        Row row = Sheet.getRow(SORY_rowindex-1);    
        Cell cell = row.getCell(1); 
        if(!(cell.getStringCellValue().equals("Operation Yield Comparison")))
            {
             correct=false;
            }    
        }
        return correct;    
    }
    
    public boolean verify_EORY_metadata_posiition( XSSFSheet Sheet,ArrayList<TokenModel> tokenmodel)
    {   
    int EORY_rowindex = -1;
    boolean correct=true;
    for(int i=0;i<tokenmodel.size();i++)
      {
         if((tokenmodel.get(i).token_name.equals("EORY")))
                 EORY_rowindex=tokenmodel.get(i).row_no;                 
      } 
    if(EORY_rowindex == -1)
    {correct=false;}
    else
    {
        Row row = Sheet.getRow(EORY_rowindex-1);    
        Cell cell = row.getCell(1); 
        if(Sheet.getLastRowNum()==EORY_rowindex-1)
        {
            correct=true;
        }
        else if(!((cell.getCellType()!=CELL_TYPE_BLANK)&&
        ((Sheet.getRow(EORY_rowindex).getCell(0)==null)||(Sheet.getRow(EORY_rowindex).getCell(0).getCellType()==CELL_TYPE_BLANK))&&
        ((Sheet.getRow(EORY_rowindex).getCell(1)==null)||(Sheet.getRow(EORY_rowindex).getCell(1).getCellType()==CELL_TYPE_BLANK))))
            {
             correct=false; 
            }
    }
    return correct;
    }
    
}
