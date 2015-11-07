/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package in.expertsoftware.colorcheck;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dharam
 */
public class VerifyWorkbook {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
       XSSFWorkbook workbook = null; 
       FileInputStream DIMT_Sheet = new FileInputStream(new File(args[0]));                   
            try {workbook = new XSSFWorkbook(DIMT_Sheet);} catch (IOException ex) 
                {    ex.printStackTrace();    }    
       VerifyTokens verifytokens=new VerifyTokens();
       ArrayList get_List=verifytokens.start(workbook);
       ArrayList<ErrorModel> get_errormodelList=(ArrayList) get_List.get(0);
       ArrayList<TokenModel> get_tokenmodelList=(ArrayList) get_List.get(1);      
       
      ///////////////// /*verify the metadata poaition of Reporting_Qtr*///////////////////
       Metadata_Position_Of_Reporting_Qtr metadataposition_qtr=new Metadata_Position_Of_Reporting_Qtr();
       int Reporting_Qtr_index=workbook.getSheetIndex("Reporting_Qtr");
       if(!(metadataposition_qtr.verify_SORQ_metadata_posiition(workbook.getSheetAt(Reporting_Qtr_index),get_tokenmodelList)))
        {
        ErrorModel errmodel=new ErrorModel();        
        errmodel.setError_desc("SORQ Token location is not correct");
        errmodel.setSheet_name("Reporting_Qtr");
        errmodel.setError_level("Error");
        get_errormodelList.add(errmodel);
        }
       if(!(metadataposition_qtr.verify_EORQ_metadata_posiition(workbook.getSheetAt(Reporting_Qtr_index),get_tokenmodelList)))
        {
        ErrorModel errmodel=new ErrorModel();        
        errmodel.setError_desc("EORQ Token location is not correct");
        errmodel.setSheet_name("Reporting_Qtr");
        errmodel.setError_level("Error");
        get_errormodelList.add(errmodel); 
        }
       
      /////////////////////// /*verify the metadata poaition of Reporting_Year*///////////////////
       Metadata_Position_Of_Reporting_Year metadataposition_year=new Metadata_Position_Of_Reporting_Year();
       int Reporting_Year_index=workbook.getSheetIndex("Reporting_Year");
       if(!(metadataposition_year.verify_SORY_metadata_posiition(workbook.getSheetAt(Reporting_Year_index),get_tokenmodelList)))
        {
        ErrorModel errmodel=new ErrorModel();        
        errmodel.setError_desc("SORY Token location is not correct");
        errmodel.setSheet_name("Reporting_Year");
        errmodel.setError_level("Error");
        get_errormodelList.add(errmodel);
        }
       if(!(metadataposition_year.verify_EORY_metadata_posiition(workbook.getSheetAt(Reporting_Year_index),get_tokenmodelList)))
        {
        ErrorModel errmodel=new ErrorModel();        
        errmodel.setError_desc("EORY Token location is not correct");
        errmodel.setSheet_name("Reporting_Year");
        errmodel.setError_level("Error");
        get_errormodelList.add(errmodel); 
        }
       int SORQtokenRow=0; 
       int EORQtokenRow=0;
       int SOOWDtokenRow=0;
       int EOOWDtokenRow=0;
       int SOFWDtokenRow=0;
       int EOFWDtokenRow=0;
       for(int i=0;i<get_tokenmodelList.size();i++)
        {
         System.out.print(get_tokenmodelList.get(i).token_name+ "row index=" + get_tokenmodelList.get(i).row_no+"\n");          
               if(get_tokenmodelList.get(i).token_name.equals("SORQ"))
                        SORQtokenRow=(get_tokenmodelList.get(i).row_no);
               else if(get_tokenmodelList.get(i).token_name.equals("EORQ"))
                        EORQtokenRow=(get_tokenmodelList.get(i).row_no);
               else if(get_tokenmodelList.get(i).token_name.equals("SOOWD"))
                        SOOWDtokenRow=(get_tokenmodelList.get(i).row_no);
               else if(get_tokenmodelList.get(i).token_name.equals("EOOWD"))
                        EOOWDtokenRow=(get_tokenmodelList.get(i).row_no);
               else if(get_tokenmodelList.get(i).token_name.equals("SOFWD"))
                        SOFWDtokenRow=(get_tokenmodelList.get(i).row_no);
               else if(get_tokenmodelList.get(i).token_name.equals("EOFWD"))
                        EOFWDtokenRow=(get_tokenmodelList.get(i).row_no);
        }
       ArrayList<String>opeartion_standard_workingSectionList=new Operation_Standard_WorkingSection_count().working_Section(SOOWDtokenRow, EOOWDtokenRow,workbook);
       ArrayList<String>financial_standard_workingSectionList=new Financial_Standard_WorkingSection_count().working_Section(SOFWDtokenRow, EOFWDtokenRow,workbook);
       ArrayList<ErrorModel> ls=(new Reporting_Qtr_Verification()).startReporting_QtrVerification(SORQtokenRow,EORQtokenRow,opeartion_standard_workingSectionList,financial_standard_workingSectionList ,workbook);
       ls.stream().forEach((errormodel) -> {
             System.out.println("ccelref "+errormodel.cell_ref+" sheetname "+errormodel.sheet_name+" dis "+errormodel.error_desc);
           });  
       ///error infomation on log
       get_errormodelList.stream().forEach((errormodel) -> {
           if(errormodel.row==-1)
           {System.out.println(errormodel.error_desc+" On sheet "+errormodel.sheet_name);}
           else if(errormodel.row==-2)
           {System.out.println(errormodel.error_desc+errormodel.sheet_name);}
           else
           {
           System.out.println("In "+errormodel.sheet_name+errormodel.error_desc+" at row"+errormodel.row+" and at colum"+errormodel.col);
           }
       });
       //finally dump the error report to the Exxcel file 
       new FormatvarificationErrorList().dumpFormatErrorToExcelFile(ls);    
   } 
 }




