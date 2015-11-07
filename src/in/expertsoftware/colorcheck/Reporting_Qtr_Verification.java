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
public class Reporting_Qtr_Verification {
    public ArrayList<ErrorModel> startReporting_QtrVerification(int SORQ_tokenLocation,int EORQ_tokenLocation,ArrayList<String>opeartion_standard_workingSectionList,ArrayList<String>financial_standard_workingSectionList,XSSFWorkbook workbook)
    {
        ArrayList<ErrorModel> errorModelList = new ArrayList<ErrorModel>();
        //start Verification
        //step first find row value of first Financial Comparision
        int first_Occurance_Of_Financial_Comparision=0;
        XSSFSheet Sheet = workbook.getSheet("Reporting_Qtr");
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
                        if((key.equals("Financial Comparison"))&&(first_Occurance_Of_Financial_Comparision==0))
                        {
                         first_Occurance_Of_Financial_Comparision= row.getRowNum()+1;
                        }
                    } 
                    }
                }
         operation_Standard_Workingsection_Verification(SORQ_tokenLocation,first_Occurance_Of_Financial_Comparision,Sheet,opeartion_standard_workingSectionList,errorModelList,workbook);               
         financial_Standard_Workingsection_Verification(first_Occurance_Of_Financial_Comparision,EORQ_tokenLocation,Sheet,financial_standard_workingSectionList,errorModelList,workbook);                                
         return errorModelList;
    }

    private void operation_Standard_Workingsection_Verification(int SORQ_tokenLocation, int first_Occurance_Of_Financial_Comparision, XSSFSheet Sheet,ArrayList<String>opeartion_standard_workingSectionList,ArrayList<ErrorModel> errorModelList,XSSFWorkbook workbook) {                   
        int reporting_Qtr_OS_Working_SectionCount=0;
        String reporting_Qtr_Formula_Cell_Formula;
        String reporting_Qtr_Formula_Cell_Value;        
        XSSFRow row;
        XSSFCell cell;        
        for(int start=(SORQ_tokenLocation-1);start<(first_Occurance_Of_Financial_Comparision-1);start++)
            {
                try
                {
                row = Sheet.getRow(start);
                cell = row.getCell(1);
                reporting_Qtr_Formula_Cell_Value=cell.getStringCellValue(); 
                if(reporting_Qtr_Formula_Cell_Value.equals("Operational Comparison"))
                {  start=start+1;   }
                else if(reporting_Qtr_OS_Working_SectionCount<opeartion_standard_workingSectionList.size())
                {       
                        switch (cell.getCellType()) 
                        {
                        case Cell.CELL_TYPE_FORMULA:                    
                             reporting_Qtr_Formula_Cell_Formula=cell.getCellFormula();
                             String retrive_Formula=opeartion_standard_workingSectionList.get(reporting_Qtr_OS_Working_SectionCount);
                             if(reporting_Qtr_Formula_Cell_Formula.equals(retrive_Formula))
                             { reporting_Qtr_OS_Working_SectionCount++;  }
                             else
                             {
                                     ErrorModel errorModel=new ErrorModel();
                                     CellReference cellRef=new CellReference(cell);
                                     errorModel.setCell_ref(cellRef.formatAsString());                             
                                     errorModel.setSheet_name("Reporting_Qtr");
                                     errorModel.setError_desc("Sequence does not match Actual Sequence should be"+retrive_Formula);
                                     errorModel.setError_level("Error");
                                     errorModelList.add(errorModel);
                                     reporting_Qtr_OS_Working_SectionCount++;
                             }                     
                            break;
                        case Cell.CELL_TYPE_BLANK:                    
                            break;                
                        default:   {                 
                                     ErrorModel errorModel=new ErrorModel();
                                     CellReference cellRef=new CellReference(cell);
                                     errorModel.setCell_ref(cellRef.formatAsString());                             
                                     errorModel.setSheet_name("Reporting_Qtr");
                                     errorModel.setError_desc("Reporting_Qtr cell does not contain formula");
                                     errorModel.setError_level("Error");
                                     errorModelList.add(errorModel);
                                     reporting_Qtr_OS_Working_SectionCount++;
                                    }
                            break;
                        }   
                }
               else
                {   reporting_Qtr_OS_Working_SectionCount++;  }
            }
            catch(NullPointerException nullexcp){continue;}
            catch(Exception e){e.printStackTrace();}       
        }           
        if(reporting_Qtr_OS_Working_SectionCount!=opeartion_standard_workingSectionList.size())
                     {
                      ErrorModel errorModel=new ErrorModel();   
                      errorModel.setSheet_name("Reporting_Qtr");
                      errorModel.setError_desc("Reporting_Qtr have "+Math.abs(reporting_Qtr_OS_Working_SectionCount-opeartion_standard_workingSectionList.size())+" extra rows from Operation_Standard");
                      errorModel.setError_level("Error");
                      errorModelList.add(errorModel);
                     }
            }
    
    private void financial_Standard_Workingsection_Verification(int first_Occurance_Of_Financial_Comparision, int EORQ_tokenLocation, XSSFSheet Sheet,ArrayList<String>financial_standard_workingSectionList,ArrayList<ErrorModel> errorModelList,XSSFWorkbook workbook) { 
        int reporting_Qtr_FS_Working_SectionCount=0;
        String reporting_Qtr_Formula_Cell_Formula = null;
        String reporting_Qtr_Formula_Cell_Value;
        XSSFRow row;
        XSSFCell cell;      
        for(int start=(first_Occurance_Of_Financial_Comparision-1);start<EORQ_tokenLocation;start++)
        { 
            try
            {   
                row = Sheet.getRow(start);
                cell = row.getCell(1);
                reporting_Qtr_Formula_Cell_Value=cell.getStringCellValue();
                if(reporting_Qtr_Formula_Cell_Value.equals("Financial Comparison"))
                    { start=start+3; }
                else if(reporting_Qtr_FS_Working_SectionCount<financial_standard_workingSectionList.size())
                {
                    switch (cell.getCellType()) 
                    {
                        case Cell.CELL_TYPE_FORMULA:                    
                             reporting_Qtr_Formula_Cell_Formula=cell.getCellFormula();                             
                             String retrive_Formula = financial_standard_workingSectionList.get(reporting_Qtr_FS_Working_SectionCount);
                             if(reporting_Qtr_Formula_Cell_Formula.equals(retrive_Formula))
                             { reporting_Qtr_FS_Working_SectionCount++;  }
                             else
                             {
                                   ErrorModel errorModel=new ErrorModel();
                                   CellReference cellRef=new CellReference(cell);
                                   errorModel.setCell_ref(cellRef.formatAsString());                             
                                   errorModel.setSheet_name("Reporting_Qtr");
                                   errorModel.setError_desc("Sequence does not match Actual Sequence should be"+retrive_Formula);
                                   errorModel.setError_level("Error");
                                   errorModelList.add(errorModel);
                                   reporting_Qtr_FS_Working_SectionCount++;
                             }                     
                            break;
                        case Cell.CELL_TYPE_BLANK:                                    
                            break;                        
                        default:   {                 
                                     ErrorModel errorModel=new ErrorModel();
                                     CellReference cellRef=new CellReference(cell);
                                     errorModel.setCell_ref(cellRef.formatAsString());                             
                                     errorModel.setSheet_name("Reporting_Qtr");
                                     errorModel.setError_desc("Reporting_Qtr cell does not contain formula ");
                                     errorModel.setError_level("Error");
                                     errorModelList.add(errorModel);
                                     reporting_Qtr_FS_Working_SectionCount++;
                                   }
                            break;
                        }
            }
            else
                 {   reporting_Qtr_FS_Working_SectionCount++; }        
           }
            catch(NullPointerException nullexcp){continue;}
            catch(Exception e){e.printStackTrace();}
         }              
        if(reporting_Qtr_FS_Working_SectionCount!=financial_standard_workingSectionList.size())
                     {
                      ErrorModel errorModel=new ErrorModel();   
                      errorModel.setSheet_name("Reporting_Qtr");
                      errorModel.setError_desc("Reporting_Qtr have "+Math.abs(reporting_Qtr_FS_Working_SectionCount-financial_standard_workingSectionList.size())+" extra rows from Operation_Standard");
                      errorModel.setError_level("Error");
                      errorModelList.add(errorModel);
                     }  
   }
}

