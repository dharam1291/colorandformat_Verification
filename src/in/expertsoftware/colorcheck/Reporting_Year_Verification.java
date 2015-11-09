/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package in.expertsoftware.colorcheck;

import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dharam
 */
public class Reporting_Year_Verification {
   public ArrayList<ErrorModel> startReporting_YearVerification(int SORY_tokenLocation,int EORY_tokenLocation,ArrayList<String>opeartion_standard_workingSectionList,ArrayList<String>financial_standard_workingSectionList,XSSFWorkbook workbook)
    {
        ArrayList<ErrorModel> errorModelList = new ArrayList<ErrorModel>();
        boolean operationCheck;
        boolean financialCheck;
        int first_Occurance_Of_Financial_Comparision=0;
        XSSFSheet Sheet = workbook.getSheet("Reporting_Year");
        Iterator<Row> rowIterator = Sheet.iterator();
                while (rowIterator.hasNext()) 
                {
                    Row row = rowIterator.next();  
                    Iterator<Cell> cellIterator = row.iterator();
                    while(cellIterator.hasNext())
                    {
                    Cell cell = cellIterator.next();
                    if(cell.getColumnIndex() == 1)
                    {
                        String key = cell.getStringCellValue();
                        if((key.equalsIgnoreCase("Financial Comparison"))&&(first_Occurance_Of_Financial_Comparision==0))
                        {
                         first_Occurance_Of_Financial_Comparision= row.getRowNum()+1;
                        }
                    } 
                    }
                } 
        operationCheck=operation_Standard_Workingsection_Verification(SORY_tokenLocation,first_Occurance_Of_Financial_Comparision,Sheet,opeartion_standard_workingSectionList,errorModelList,workbook);               
        //financialCheck=financial_Standard_Workingsection_Verification(first_Occurance_Of_Financial_Comparision,EORY_tokenLocation,Sheet,financial_standard_workingSectionList,errorModelList,workbook);                                
        
        //if return false that means no error.
        if(!operationCheck)
        {
         reporting_Year_operation_Standard_C_To_I_Column_Verification((SORY_tokenLocation+3),first_Occurance_Of_Financial_Comparision,errorModelList,Sheet,workbook);   
        }
        /*if(!financialCheck)
        {
          
        }*/
        return errorModelList;
    }

    private boolean operation_Standard_Workingsection_Verification(int SORY_tokenLocation, int first_Occurance_Of_Financial_Comparision, XSSFSheet Sheet,ArrayList<String>opeartion_standard_workingSectionList,ArrayList<ErrorModel> errorModelList,XSSFWorkbook workbook) {                   
        int reporting_Year_OS_Working_SectionCount=0;
        String reporting_Year_Formula_Cell_Formula;
        String reporting_Year_Formula_Cell_Value; 
        boolean isError=false;
        XSSFRow row;
        XSSFCell cell;        
        for(int start=(SORY_tokenLocation-1);start<(first_Occurance_Of_Financial_Comparision-1);start++)
            {
                try
                {
                row = Sheet.getRow(start);
                cell = row.getCell(1);
                reporting_Year_Formula_Cell_Value=cell.getStringCellValue();               
                if(reporting_Year_Formula_Cell_Value.equalsIgnoreCase("Operational Comparison"))
                {  start=start+1;   }
                else if(reporting_Year_OS_Working_SectionCount<opeartion_standard_workingSectionList.size())
                {       
                        switch (cell.getCellType()) 
                        {
                        case Cell.CELL_TYPE_FORMULA:                    
                             reporting_Year_Formula_Cell_Formula=cell.getCellFormula();
                             if(reporting_Year_Formula_Cell_Formula.contains("$"))
                             {
                               reporting_Year_Formula_Cell_Formula=reporting_Year_Formula_Cell_Formula.replaceAll("\\$","").replaceAll(" ", "");
                             }
                             String retrive_Formula=opeartion_standard_workingSectionList.get(reporting_Year_OS_Working_SectionCount);
                             if(reporting_Year_Formula_Cell_Formula.equals(retrive_Formula))
                             { reporting_Year_OS_Working_SectionCount++;  }
                             else
                             {
                                     ErrorModel errorModel=new ErrorModel();
                                     CellReference cellRef=new CellReference(cell);
                                     errorModel.setCell_ref(cellRef.formatAsString());                             
                                     errorModel.setSheet_name("Reporting_Year");
                                     errorModel.setError_desc("Sequence does not match Actual Sequence should be"+retrive_Formula);
                                     errorModel.setError_level("Error");
                                     errorModel.setRow((cell.getRowIndex()+1));
                                     errorModelList.add(errorModel);
                                     reporting_Year_OS_Working_SectionCount++;
                                     isError=true;
                             }                     
                            break;
                        case Cell.CELL_TYPE_BLANK:                    
                            break;                
                        default:   {                 
                                     ErrorModel errorModel=new ErrorModel();
                                     CellReference cellRef=new CellReference(cell);
                                     errorModel.setCell_ref(cellRef.formatAsString());                             
                                     errorModel.setSheet_name("Reporting_Year");
                                     errorModel.setError_desc("Cell does not contain formula");
                                     errorModel.setError_level("Error");
                                     errorModel.setRow((cell.getRowIndex()+1));
                                     errorModelList.add(errorModel);
                                     reporting_Year_OS_Working_SectionCount++;
                                     isError=true;
                                    }
                            break;
                        }   
                }
               else
                {   reporting_Year_OS_Working_SectionCount++;  }
            }
            catch(NullPointerException nullexcp){continue;}
            catch(Exception e){e.printStackTrace();}       
        }           
        if(reporting_Year_OS_Working_SectionCount!=opeartion_standard_workingSectionList.size())
                     {
                      ErrorModel errorModel=new ErrorModel();   
                      errorModel.setSheet_name("Reporting_Year");
                      errorModel.setError_desc("Reporting_Year have "+Math.abs(reporting_Year_OS_Working_SectionCount-opeartion_standard_workingSectionList.size())+" extra rows from Operation_Standard");
                      errorModel.setError_level("Error");
                      errorModelList.add(errorModel);
                      isError=true;
                     }
         return isError;
            }
    
    private boolean financial_Standard_Workingsection_Verification(int first_Occurance_Of_Financial_Comparision, int EORY_tokenLocation, XSSFSheet Sheet,ArrayList<String>financial_standard_workingSectionList,ArrayList<ErrorModel> errorModelList,XSSFWorkbook workbook) { 
        int reporting_Year_FS_Working_SectionCount=0;
        String reporting_Year_Formula_Cell_Formula = null;
        String reporting_Year_Formula_Cell_Value;
        XSSFRow row;
        XSSFCell cell;
        boolean isError=false;
        for(int start=(first_Occurance_Of_Financial_Comparision-1);start<EORY_tokenLocation;start++)
        { 
            try
            {   
                row = Sheet.getRow(start);
                cell = row.getCell(1);
                reporting_Year_Formula_Cell_Value=cell.getStringCellValue();
                if(reporting_Year_Formula_Cell_Value.equalsIgnoreCase("Financial Comparison"))
                    { start=start+3; }
                else if(reporting_Year_FS_Working_SectionCount<financial_standard_workingSectionList.size())
                {
                    switch (cell.getCellType()) 
                    {
                        case Cell.CELL_TYPE_FORMULA:                    
                             reporting_Year_Formula_Cell_Formula=cell.getCellFormula();                             
                             if(reporting_Year_Formula_Cell_Formula.contains("$"))
                             {
                               reporting_Year_Formula_Cell_Formula=reporting_Year_Formula_Cell_Formula.replaceAll("\\$","").replaceAll(" ", "");
                             }
                             String retrive_Formula = financial_standard_workingSectionList.get(reporting_Year_FS_Working_SectionCount);
                             if(reporting_Year_Formula_Cell_Formula.equals(retrive_Formula))
                             { reporting_Year_FS_Working_SectionCount++;  }
                             else
                             {
                                   ErrorModel errorModel=new ErrorModel();
                                   CellReference cellRef=new CellReference(cell);
                                   errorModel.setCell_ref(cellRef.formatAsString());                             
                                   errorModel.setSheet_name("Reporting_Year");
                                   errorModel.setError_desc("Sequence does not match Actual Sequence should be"+retrive_Formula);
                                   errorModel.setError_level("Error");
                                   errorModel.setRow((cell.getRowIndex()+1));
                                   errorModelList.add(errorModel);
                                   reporting_Year_FS_Working_SectionCount++;
                                   isError=true;
                             }                     
                            break;
                        case Cell.CELL_TYPE_BLANK:                                    
                            break;                        
                        default:   {                 
                                     ErrorModel errorModel=new ErrorModel();
                                     CellReference cellRef=new CellReference(cell);
                                     errorModel.setCell_ref(cellRef.formatAsString());                             
                                     errorModel.setSheet_name("Reporting_Year");
                                     errorModel.setError_desc("Cell does not contain formula ");
                                     errorModel.setError_level("Error");
                                     errorModel.setRow((cell.getRowIndex()+1));
                                     errorModelList.add(errorModel);
                                     reporting_Year_FS_Working_SectionCount++;
                                     isError=true;
                                   }
                            break;
                        }
            }
            else
                 {   reporting_Year_FS_Working_SectionCount++; }        
           }
            catch(NullPointerException nullexcp){continue;}
            catch(Exception e){e.printStackTrace();}
         }              
        if(reporting_Year_FS_Working_SectionCount!=financial_standard_workingSectionList.size())
                     {
                      ErrorModel errorModel=new ErrorModel();   
                      errorModel.setSheet_name("Reporting_Year");
                      errorModel.setError_desc("Reporting_Year have "+Math.abs(reporting_Year_FS_Working_SectionCount-financial_standard_workingSectionList.size())+" extra rows from Operation_Standard");
                      errorModel.setError_level("Error");
                      errorModelList.add(errorModel);
                      isError=true;
                     }  
        return isError;
   }

    private void reporting_Year_operation_Standard_C_To_I_Column_Verification(int start_Point, int first_Occurance_Of_Financial_Comparision, ArrayList<ErrorModel> errorModelList, XSSFSheet Sheet, XSSFWorkbook workbook) {
    String reporting_Year_Formula_Cell_Formula = null;
    String reporting_Year_Formula_Cell_Value;
    String reporting_Year_B_Column_Formula=null;
    XSSFRow row;
    XSSFCell cell_B,cell_C,cell_E = null,cell_F = null,cell_G = null,cell_H = null,cell_I = null;
     for(int start=(start_Point-1);start<(first_Occurance_Of_Financial_Comparision-1);start++)
            {
                try
                {
                row = Sheet.getRow(start);
                cell_B=row.getCell(1);
                cell_C = row.getCell(2);
                cell_E = row.getCell(4);
                cell_F = row.getCell(5);
                cell_G = row.getCell(6);
                cell_H = row.getCell(7);
                cell_I = row.getCell(8);
                switch (cell_C.getCellType()) 
                        {
                        case Cell.CELL_TYPE_FORMULA:                    
                             reporting_Year_Formula_Cell_Formula=cell_C.getCellFormula();
                             if(reporting_Year_Formula_Cell_Formula.contains("$"))
                             {  reporting_Year_Formula_Cell_Formula=reporting_Year_Formula_Cell_Formula.replaceAll("\\$","").replaceAll(" ", "");   }                                                                                                            
               
                            //verify the formula is correct or not.
                             if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Year_Formula_Cell_Formula.charAt(19)=='D'))
                               { 
                                   if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Year_Formula_Cell_Formula.substring(20, reporting_Year_Formula_Cell_Formula.length()))))
                                   {    genrateError(cell_C,errorModelList,"Operation_Standard");    }
                               } 
                             //else throw an error
                             else
                             {  genrateError(cell_C,errorModelList,"Operation_Standard");   } 
                             verify_E(cell_E,cell_B,errorModelList);
                             verify_F(cell_F,cell_B,errorModelList);
                             verify_G(cell_G,cell_B,errorModelList);
                             verify_H(cell_H,cell_B,errorModelList);
                             verify_I(cell_I,cell_B,errorModelList);
                             break;
                        case Cell.CELL_TYPE_STRING:
                            reporting_Year_Formula_Cell_Value=cell_C.getStringCellValue();
                            if(reporting_Year_Formula_Cell_Value.contains("USD"))
                            {                          
                            
                            }
                            else if (reporting_Year_Formula_Cell_Value.equalsIgnoreCase("Unit"))
                            { start=start+4;}
                            else
                            {   genrateError(cell_C,errorModelList,"Operation_Standard");    }
                            break;
                        case Cell.CELL_TYPE_BLANK:                            
                            break;
                        default:
                                   genrateError(cell_C,errorModelList,"Operation_Standard");        
                            break;
                       }
                } 
            catch(NullPointerException nullexcp){continue;}
            catch(Exception e){e.printStackTrace();}       
            }
        }

    private void verify_E(XSSFCell cell_E,XSSFCell cell_B,ArrayList<ErrorModel> errorModelList) {
        String reporting_Year_Formula_Cell_Formula;
                switch(cell_E.getCellType())
                 {
                case Cell.CELL_TYPE_FORMULA:
                reporting_Year_Formula_Cell_Formula=cell_E.getCellFormula();
                if(reporting_Year_Formula_Cell_Formula.contains("$"))
                    {  reporting_Year_Formula_Cell_Formula=reporting_Year_Formula_Cell_Formula.replaceAll("\\$","").replaceAll(" ", "");   }  
                    //verify the formula
                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Year_Formula_Cell_Formula.charAt(19)=='K')&&(!(reporting_Year_Formula_Cell_Formula.contains("/"))))
                 {                  
                if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Year_Formula_Cell_Formula.substring(20, reporting_Year_Formula_Cell_Formula.length()))))
                  { genrateError(cell_E,errorModelList,"Operation_Standard");   }
                 }
             //else throw an error
             else
               {   
                   genrateError(cell_E,errorModelList,"Operation_Standard");    }                                  
              break;
          case Cell.CELL_TYPE_BLANK:
              break;
          default:
               genrateError(cell_E,errorModelList,"Operation_Standard"); 
           break;
         }    
       }

    private void verify_F(XSSFCell cell_F, XSSFCell cell_B, ArrayList<ErrorModel> errorModelList) {
        String reporting_Year_Formula_Cell_Formula;
                switch(cell_F.getCellType())
                 {
                case Cell.CELL_TYPE_FORMULA:
                reporting_Year_Formula_Cell_Formula=cell_F.getCellFormula();
                if(reporting_Year_Formula_Cell_Formula.contains("$"))
                    {  reporting_Year_Formula_Cell_Formula=reporting_Year_Formula_Cell_Formula.replaceAll("\\$","").replaceAll(" ", "");   }  
                    //verify the formula
                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Year_Formula_Cell_Formula.charAt(19)=='L')&&(!(reporting_Year_Formula_Cell_Formula.contains("/"))))
                 {                  
                if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Year_Formula_Cell_Formula.substring(20, reporting_Year_Formula_Cell_Formula.length()))))
                  { genrateError(cell_F,errorModelList,"Operation_Standard");   }
                 }
             //else throw an error
             else
               {   
                   genrateError(cell_F,errorModelList,"Operation_Standard");    }                                  
              break;
          case Cell.CELL_TYPE_BLANK:
              break;
          default:
               genrateError(cell_F,errorModelList,"Operation_Standard"); 
           break;
         }    
        }

    private void verify_G(XSSFCell cell_G, XSSFCell cell_B, ArrayList<ErrorModel> errorModelList) {
        String reporting_Year_Formula_Cell_Formula;
                switch(cell_G.getCellType())
                 {
                case Cell.CELL_TYPE_FORMULA:
                reporting_Year_Formula_Cell_Formula=cell_G.getCellFormula();
                if(reporting_Year_Formula_Cell_Formula.contains("$"))
                    {  reporting_Year_Formula_Cell_Formula=reporting_Year_Formula_Cell_Formula.replaceAll("\\$","").replaceAll(" ", "");   }  
                    //verify the formula
                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Year_Formula_Cell_Formula.charAt(19)=='M')&&(!(reporting_Year_Formula_Cell_Formula.contains("/"))))
                 {                  
                if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Year_Formula_Cell_Formula.substring(20, reporting_Year_Formula_Cell_Formula.length()))))
                  { genrateError(cell_G,errorModelList,"Operation_Standard");   }
                 }
             //else throw an error
             else
               {   
                   genrateError(cell_G,errorModelList,"Operation_Standard");    }                                  
              break;
          case Cell.CELL_TYPE_BLANK:
              break;
          default:
               genrateError(cell_G,errorModelList,"Operation_Standard"); 
           break;
         }    
        }

    private void verify_H(XSSFCell cell_H, XSSFCell cell_B, ArrayList<ErrorModel> errorModelList) {
        String reporting_Year_Formula_Cell_Formula;
                switch(cell_H.getCellType())
                 {
                case Cell.CELL_TYPE_FORMULA:
                reporting_Year_Formula_Cell_Formula=cell_H.getCellFormula();
                if(reporting_Year_Formula_Cell_Formula.contains("$"))
                    {  reporting_Year_Formula_Cell_Formula=reporting_Year_Formula_Cell_Formula.replaceAll("\\$","").replaceAll(" ", "");   }  
                    //verify the formula
                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Year_Formula_Cell_Formula.charAt(19)=='N')&&(!(reporting_Year_Formula_Cell_Formula.contains("/"))))
                 {                  
                if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Year_Formula_Cell_Formula.substring(20, reporting_Year_Formula_Cell_Formula.length()))))
                  { genrateError(cell_H,errorModelList,"Operation_Standard");   }
                 }
             //else throw an error
             else
               {   
                   genrateError(cell_H,errorModelList,"Operation_Standard");    }                                  
              break;
          case Cell.CELL_TYPE_BLANK:
              break;
          default:
               genrateError(cell_H,errorModelList,"Operation_Standard"); 
           break;
         }    
        }

    private void verify_I(XSSFCell cell_I, XSSFCell cell_B, ArrayList<ErrorModel> errorModelList) {
        String reporting_Year_Formula_Cell_Formula;
                switch(cell_I.getCellType())
                 {
                case Cell.CELL_TYPE_FORMULA:
                reporting_Year_Formula_Cell_Formula=cell_I.getCellFormula();
                if(reporting_Year_Formula_Cell_Formula.contains("$"))
                    {  reporting_Year_Formula_Cell_Formula=reporting_Year_Formula_Cell_Formula.replaceAll("\\$","").replaceAll(" ", "");   }  
                    //verify the formula
                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Year_Formula_Cell_Formula.charAt(19)=='O')&&(!(reporting_Year_Formula_Cell_Formula.contains("/"))))
                 {                  
                if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Year_Formula_Cell_Formula.substring(20, reporting_Year_Formula_Cell_Formula.length()))))
                  { genrateError(cell_I,errorModelList,"Operation_Standard");   }
                 }
             //else throw an error
             else
               {   
                   genrateError(cell_I,errorModelList,"Operation_Standard");    }                                  
              break;
          case Cell.CELL_TYPE_BLANK:
              break;
          default:
               genrateError(cell_I,errorModelList,"Operation_Standard"); 
           break;
         }    
        }

    private void genrateError(XSSFCell cell_ref, ArrayList<ErrorModel> errorModelList,String worksection_type) {
        ErrorModel errorModel=new ErrorModel();
        CellReference cellRef=new CellReference(cell_ref);
        errorModel.setCell_ref(cellRef.formatAsString());                             
        errorModel.setSheet_name("Reporting_Year");
        errorModel.setError_desc("Cell Formula is not linked correctly to the "+worksection_type);
        errorModel.setError_level("Error");
        errorModel.setRow((cell_ref.getRowIndex()+1));
        errorModelList.add(errorModel);    
    }
   
    
    
    
}
