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
public class Metadata_Position_Of_Reporting_Qtr {
    public boolean verify_SORQ_metadata_posiition( XSSFSheet Sheet,ArrayList<TokenModel> tokenmodel)
    {    
    int SORQ_rowindex = -1;
    boolean correct=true;
    for(int i=0;i<tokenmodel.size();i++)
      {
         if((tokenmodel.get(i).token_name.equals("SORQ")))
                 SORQ_rowindex=tokenmodel.get(i).row_no;                 
      }
    if(SORQ_rowindex==-1)
    {correct=false;}
    else
    {
        Row row = Sheet.getRow(SORQ_rowindex-1);    
        Cell cell = row.getCell(1); 
        if(!(cell.getStringCellValue().equals("Operation Yield Comparison")))
            {
             correct=false;
            }
    }
    return correct;
    }
    public boolean verify_EORQ_metadata_posiition( XSSFSheet Sheet,ArrayList<TokenModel> tokenmodel)
    {   
    int EORQ_rowindex = -1;
    boolean correct=true;
    for(int i=0;i<tokenmodel.size();i++)
      {
         if((tokenmodel.get(i).token_name.equals("EORQ")))
                 EORQ_rowindex=tokenmodel.get(i).row_no;                 
      } 
    if(EORQ_rowindex==-1)
    {correct=false;}
    else
    {
        Row row = Sheet.getRow(EORQ_rowindex-1);    
        Cell cell = row.getCell(1); 
        if(Sheet.getLastRowNum()==EORQ_rowindex-1)
        {
            correct=true;
        }
        else if(!((cell.getCellType()!=CELL_TYPE_BLANK)&&
        ((Sheet.getRow(EORQ_rowindex).getCell(0)==null)||(Sheet.getRow(EORQ_rowindex).getCell(0).getCellType()==CELL_TYPE_BLANK))&&
        ((Sheet.getRow(EORQ_rowindex).getCell(1)==null)||(Sheet.getRow(EORQ_rowindex).getCell(1).getCellType()==CELL_TYPE_BLANK))))
            {
             correct=false; 
            }
    }
    return correct;
    }
}
