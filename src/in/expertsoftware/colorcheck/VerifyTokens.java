/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package in.expertsoftware.colorcheck;




import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.springframework.stereotype.Service;

/**
 *
 * @author Dharam
 */
public class VerifyTokens {    
    String[] Tokens={"SOUFI","UFCS1","EOUFI","UFCS2","SOUOI","EOUOI",
            "SOFI","FCS1","EOFI","FCS2","SOFWD","EOFWD","SOFCP","EOFCP",
            "SOOI","EOOI","SOOCS","EOOCS","SOOWD","EOOWD","SOCCK","EOCCK","SORQ","EORQ","SORY","EORY","SOOCQ","EOOCQ","SOFCQ","EOFCQ","SOOCY","EOOCY","SOFCY","EOFCY","SOS","EOS"}; 
    /**
     * 
     * Start is an initiate function of VerifyTokens 
     * @param workbook an instance of XSSFWorkbook
     * @return ArrayList of type Error and Token class
     */
    public ArrayList start( XSSFWorkbook workbook)
    { 
    	   int NumberOfSheets = workbook.getNumberOfSheets();           
           String[] SheetName = new String[NumberOfSheets];            
           System.out.println("Get Sheets from workwook");          
           /*for(int i=0;i<NumberOfSheets;i++)
           {    SheetName[i]=workbook.getSheetName(i); }  */         
           System.out.println("Tokens");
           System.out.println(Arrays.deepToString(Tokens));
           
           System.out.println("Check Tokens from Workbook");
           
           ArrayList errorAndTokenList=verify_tokens(NumberOfSheets,workbook); 
           
           return errorAndTokenList;               
    } 
/**
 * Verify_tokens function takes two parameters first is number of sheets present in workbook and second is an instance of workbook.
 * it process each sheet individually and verifying the tokens position as well check tokens are present or not and error is handled by error model.
 * @param NumberOfSheets Number of Sheet present In this workbook
 * @param workbook an instance of XSSFWorkbook
 * @return ArrayList   of type Error and Model
 * 
 */
    public ArrayList verify_tokens(int NumberOfSheets,XSSFWorkbook workbook) {
    	ArrayList<ErrorModel> errorModelList=new ArrayList<ErrorModel>();
        ArrayList<TokenModel> tokenModelList=new ArrayList<TokenModel>(); 
        ArrayList<String> metadataCount =new ArrayList<String>(); 
        ArrayList errorAndTokenList=new ArrayList();
        for(int i=0;i<NumberOfSheets;i++)
           {          
           XSSFSheet Sheet = workbook.getSheetAt(i);
           if(containSheetName(Sheet.getSheetName()))
           {
           Iterator<Row> rowIterator = Sheet.iterator();
                while (rowIterator.hasNext()) 
                {
                    Row row = rowIterator.next();  
                    Iterator<Cell> cellIterator = row.iterator();
                    while(cellIterator.hasNext())
                    {
                    Cell cell = cellIterator.next();
                    if(cell.getColumnIndex() == 0)
                    {
                        String key = cell.getStringCellValue();
                        switch(Sheet.getSheetName())
                        {
                            case "BasicInfo": if(!(cell.getCellType()==Cell.CELL_TYPE_BLANK))
                                              {
                                              ErrorModel errmodel=new ErrorModel();
                                              errmodel.setSheet_name(Sheet.getSheetName());
                                              CellReference cellRef=new CellReference(cell);
                                              errmodel.setCell_ref(cellRef.formatAsString());
                                              errmodel.setRow(cell.getRowIndex()+1);
                                              errmodel.setCol(cell.getColumnIndex()+1);
                                              errmodel.setError_desc("Vulnerable Token present at");
                                              errmodel.setError_level("Error");
                                              errorModelList.add(errmodel);
                                              }
                                              break;                                                                                     
                            case "User_Financial_Input":switch(key.trim())
                                                        {
                                                        case "SOUFI":TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOUFI"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {                                                                                                                                          
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));                                                                                                                                        
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOUFI");
                                                                     }
                                                                     break;
                                                                     
                                                        case "EOUFI":                                                                     
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOUFI"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel); 
                                                                     metadataCount.add("EOUFI");
                                                                     }
                                                                     break;
                                                            
                                                       case "UFCS1": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("UFCS1"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {                                                                                                                                          
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));                                                                                                                                        
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("UFCS1");
                                                                     }
                                                                     break;
                                                                     
                                                        case "UFCS2":                                                                     
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("UFCS2"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel); 
                                                                     metadataCount.add("UFCS2");
                                                                     }
                                                                     break;
                                                            }                                                                    
                                                    break;
                            case "User_Operation_Input": switch(key.trim())
                                                        {
                                                        
                                                        case "SOUOI":TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOUOI"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {                                                                                                                                          
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));                                                                                                                                        
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOUOI");
                                                                     }
                                                                     break;
                                                                     
                                                        case "EOUOI":                                                                     
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOUOI"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel); 
                                                                     metadataCount.add("EOUOI");
                                                                     }
                                                                     break;
                                                        }                           
                                    break;
                            case "Unit_Map":break;
                            case "Operation_Standard":  switch(key.trim())
                                                        {
                                                            case "SOOI": 
                                                                        TokenModel tknmodel=new TokenModel();
                                                                        if(metadataCount.contains("SOOI"))
                                                                        {
                                                                        ErrorModel errmodel=new ErrorModel();
                                                                        errmodel.setSheet_name(Sheet.getSheetName());
                                                                        CellReference cellRef=new CellReference(cell);
                                                                        errmodel.setCell_ref(cellRef.formatAsString());
                                                                        errmodel.setRow(cell.getRowIndex()+1);
                                                                        errmodel.setCol(cell.getColumnIndex()+1);
                                                                        errmodel.setError_desc("Token Present More Than One Time");
                                                                        errmodel.setError_level("Error");
                                                                        errorModelList.add(errmodel);
                                                                        }
                                                                        else
                                                                        {
                                                                         tknmodel.setSheet_name(Sheet.getSheetName());
                                                                         tknmodel.setToken_name(key.trim());
                                                                         tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                         tokenModelList.add(tknmodel);
                                                                         metadataCount.add("SOOI");
                                                                        }
                                                                  break;
                                                            case "EOOI": 
                                                                         tknmodel=new TokenModel();
                                                                         if(metadataCount.contains("EOOI"))
                                                                         {
                                                                         ErrorModel errmodel=new ErrorModel();
                                                                         CellReference cellRef=new CellReference(cell);
                                                                         errmodel.setSheet_name(Sheet.getSheetName());
                                                                         CellReference cellRef1=new CellReference(cell);
                                                                         errmodel.setCell_ref(cellRef1.formatAsString());
                                                                         errmodel.setRow(cell.getRowIndex()+1);
                                                                         errmodel.setCol(cell.getColumnIndex()+1);
                                                                         errmodel.setError_desc("Token Present More Than One Time");
                                                                         errmodel.setError_level("Error");
                                                                         errorModelList.add(errmodel);
                                                                         }
                                                                         else
                                                                         {
                                                                         tknmodel.setSheet_name(Sheet.getSheetName());
                                                                         tknmodel.setToken_name(key.trim());
                                                                         tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                         tokenModelList.add(tknmodel);
                                                                         metadataCount.add("EOOI");
                                                                         }
                                                                 break;
                                                            case "SOOCS": 
                                                                         tknmodel=new TokenModel();
                                                                         if(metadataCount.contains("SOOCS"))
                                                                         {
                                                                         ErrorModel errmodel=new ErrorModel();
                                                                         errmodel.setSheet_name(Sheet.getSheetName());
                                                                         CellReference cellRef=new CellReference(cell);
                                                                         errmodel.setCell_ref(cellRef.formatAsString());
                                                                         errmodel.setRow(cell.getRowIndex()+1);
                                                                         errmodel.setCol(cell.getColumnIndex()+1);
                                                                         errmodel.setError_desc("Token Present More Than One Time");
                                                                         errmodel.setError_level("Error");
                                                                         errorModelList.add(errmodel);
                                                                         }
                                                                         else
                                                                         {
                                                                         tknmodel.setSheet_name(Sheet.getSheetName());
                                                                         tknmodel.setToken_name(key.trim());
                                                                         tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                         tokenModelList.add(tknmodel);
                                                                         metadataCount.add("SOOCS");
                                                                         }
                                                                  break;
                                                            case "EOOCS": 
                                                                         tknmodel=new TokenModel();
                                                                         if(metadataCount.contains("EOOCS"))
                                                                         {
                                                                         ErrorModel errmodel=new ErrorModel();
                                                                         errmodel.setSheet_name(Sheet.getSheetName());
                                                                         CellReference cellRef=new CellReference(cell);
                                                                         errmodel.setCell_ref(cellRef.formatAsString());
                                                                         errmodel.setRow(cell.getRowIndex()+1);
                                                                         errmodel.setCol(cell.getColumnIndex()+1);
                                                                         errmodel.setError_desc("Token Present More Than One Time");
                                                                         errmodel.setError_level("Error");
                                                                         errorModelList.add(errmodel);
                                                                         }
                                                                         else
                                                                         {
                                                                         tknmodel.setSheet_name(Sheet.getSheetName());
                                                                         tknmodel.setToken_name(key.trim());
                                                                         tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                         tokenModelList.add(tknmodel);
                                                                         metadataCount.add("EOOCS");
                                                                         }                                       
                                                                 break;
                                                            case "SOOWD": 
                                                                         tknmodel=new TokenModel();
                                                                         if(metadataCount.contains("SOOWD"))
                                                                         {
                                                                         ErrorModel errmodel=new ErrorModel();
                                                                         errmodel.setSheet_name(Sheet.getSheetName());
                                                                         CellReference cellRef=new CellReference(cell);
                                                                         errmodel.setCell_ref(cellRef.formatAsString());
                                                                         errmodel.setRow(cell.getRowIndex()+1);
                                                                         errmodel.setCol(cell.getColumnIndex()+1);
                                                                         errmodel.setError_desc("Token Present More Than One Time");
                                                                         errmodel.setError_level("Error");
                                                                         errorModelList.add(errmodel);
                                                                         }
                                                                         else
                                                                         {
                                                                         tknmodel.setSheet_name(Sheet.getSheetName());
                                                                         tknmodel.setToken_name(key.trim());
                                                                         tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                         tokenModelList.add(tknmodel);
                                                                         metadataCount.add("SOOWD");
                                                                         }
                                                                  break;
                                                            case "EOOWD": 
                                                                         tknmodel=new TokenModel();
                                                                         if(metadataCount.contains("EOOWD"))
                                                                         {
                                                                         ErrorModel errmodel=new ErrorModel();
                                                                         errmodel.setSheet_name(Sheet.getSheetName());
                                                                         CellReference cellRef=new CellReference(cell);
                                                                         errmodel.setCell_ref(cellRef.formatAsString());
                                                                         errmodel.setRow(cell.getRowIndex()+1);
                                                                         errmodel.setCol(cell.getColumnIndex()+1);
                                                                         errmodel.setError_desc("Token Present More Than One Time");
                                                                         errmodel.setError_level("Error");
                                                                         errorModelList.add(errmodel);
                                                                         }
                                                                         else
                                                                         {
                                                                         tknmodel.setSheet_name(Sheet.getSheetName());
                                                                         tknmodel.setToken_name(key.trim());
                                                                         tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                         tokenModelList.add(tknmodel);
                                                                         metadataCount.add("EOOWD");
                                                                         }
                                                                 break;
                                                            
                                                        }
                                        break; 
                            case "Financial_Standard": switch(key.trim())
                                                    {
                                                        case "SOFI": 
                                                                     TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOFI"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOFI");
                                                                     }
                                                              break;
                                                        case "EOFI": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOFI"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EOFI");
                                                                     }
                                                             break;
                                                        case "FCS1": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("FCS1"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {    
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("FCS1");
                                                                     }
                                                              break;
                                                        case "FCS2": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("FCS2"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("FCS2");
                                                                     }
                                                             break;
                                                        case "SOFWD": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOFWD"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOFWD");
                                                                     }
                                                              break;
                                                        case "EOFWD": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOFWD"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EOFWD");
                                                                     }
                                                             break;
                                                        case "SOFCP": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOFCP"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOFCP");
                                                                     }
                                                              break;
                                                        case "EOFCP": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOFCP"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EOFCP");
                                                                     }
                                                             break;
                                                    }
                                        break; 
                            case "CrossCheck": switch(key.trim())
                                                {
                                                    case "SOCCK": 
                                                                     TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOCCK"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {    
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOCCK");
                                                                     }
                                                          break;
                                                    case "EOCCK": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOCCK"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EOCCK");
                                                                     }
                                                         break;
                                                }
                                        break; 
                            case "Reporting_Qtr": switch(key.trim())
                                                    {
                                                        case "SORQ": 
                                                                     TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SORQ"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SORQ");
                                                                     }
                                                              break;
                                                        case "EORQ": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EORQ"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EORQ");
                                                                     }
                                                             break;
                                                    }
                                        break; 
                            case "Reporting_Year": switch(key.trim())
                                                    {
                                                        case "SORY": 
                                                                     TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SORY"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SORY");
                                                                     }
                                                              break;
                                                        case "EORY": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EORY"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EORY");
                                                                     }
                                                             break;
                                                    }
                                        break; 
                            case "Chart_Qtr": switch(key.trim())
                                                {
                                                    case "SOOCQ": 
                                                                     TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOOCQ"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOOCQ");
                                                                     }
                                                          break;
                                                    case "EOOCQ": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOOCQ"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EOOCQ");
                                                                     }
                                                         break;
                                                    case "SOFCQ": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOFCQ"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOFCQ");
                                                                     }
                                                          break;
                                                    case "EOFCQ": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOFCQ"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);                                                                    
                                                                     metadataCount.add("EOFCQ");
                                                                     }
                                                         break;
                                                }
                                        break; 
                            case "Chart_Year": switch(key.trim())
                                                {
                                                    case "SOOCY": 
                                                                     TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOOCY"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOOCY");
                                                                     }
                                                          break;
                                                    case "EOOCY": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOOCY"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EOOCY");
                                                                     }
                                                         break;
                                                    case "SOFCY": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOFCY"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOFCY");
                                                                     }
                                                          break;
                                                    case "EOFCY": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOFCY"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EOFCY");
                                                                     }
                                                         break;
                                                }
                                        break; 
                            case "Summary": switch(key.trim())
                                            {
                                                    case "SOS": 
                                                                     TokenModel tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("SOS"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("SOS");
                                                                     }
                                                      break;
                                                    case "EOS": 
                                                                     tknmodel=new TokenModel();
                                                                     if(metadataCount.contains("EOS"))
                                                                     {
                                                                     ErrorModel errmodel=new ErrorModel();
                                                                     errmodel.setSheet_name(Sheet.getSheetName());
                                                                     CellReference cellRef=new CellReference(cell);
                                                                     errmodel.setCell_ref(cellRef.formatAsString());
                                                                     errmodel.setRow(cell.getRowIndex()+1);
                                                                     errmodel.setCol(cell.getColumnIndex()+1);
                                                                     errmodel.setError_desc("Token Present More Than One Time");
                                                                     errmodel.setError_level("Error");
                                                                     errorModelList.add(errmodel);
                                                                     }
                                                                     else
                                                                     {
                                                                     tknmodel.setSheet_name(Sheet.getSheetName());
                                                                     tknmodel.setToken_name(key.trim());
                                                                     tknmodel.setRow_no((cell.getRowIndex()+1));  
                                                                     tokenModelList.add(tknmodel);
                                                                     metadataCount.add("EOS");
                                                                     }
                                                      break;
                                            }
                                        break; 
                            case "MetaDataSheet":break;                           
                        }
                    } 
                  }
               }
           }
           else
            {
            System.out.println("workbook have another sheet");
            ErrorModel errmodel=new ErrorModel();
            errmodel.setError_desc("This is an extra sheet in this workbook");
            errmodel.setSheet_name(Sheet.getSheetName());
            //errmodel.setRow(-2);
            errmodel.setError_level("Warning");
            errorModelList.add(errmodel);   
            }
        }
        //for checking every tocken is present or not;  
        errorModelList.addAll(verifyAllTokenPresent(metadataCount));
        errorAndTokenList.add(errorModelList);
        errorAndTokenList.add(tokenModelList);
        return errorAndTokenList;
    }        

    /**
     * 
     * @param metadatacount a specific token name
     * @return name of sheet corresponding to token name
     */
    private ArrayList<ErrorModel> verifyAllTokenPresent(ArrayList<String> metadatacount) {
    	ArrayList<ErrorModel> errorModelList=new ArrayList<ErrorModel>();
        for (int i = 0; i < Tokens.length; i++) {
            if (!(metadatacount.contains(Tokens[i]))) {
                String sheetname = null;
                if(Tokens[i].equals("SOUOI")||Tokens[i].equals("EOUOI"))
                    sheetname="User_Operation_input";
                else if(Tokens[i].equals("SOOI")||Tokens[i].equals("EOOI")||Tokens[i].equals("SOOCS")||Tokens[i].equals("EOOCS")||Tokens[i].equals("SOOWD")||Tokens[i].equals("EOOWD"))
                    sheetname="Operation_Standard";
                else if(Tokens[i].equals("SOFI")||Tokens[i].equals("EOFI")||Tokens[i].equals("FCS1")||Tokens[i].equals("FCS2")||Tokens[i].equals("SOFWD")||Tokens[i].equals("EOFWD")||Tokens[i].equals("SOFCP")||Tokens[i].equals("EOFCP"))
                    sheetname="Financial_Standard";
                else if(Tokens[i].equals("SOUFI")||Tokens[i].equals("EOUFI")||Tokens[i].equals("UFCS1")||Tokens[i].equals("UFCS2"))
                    sheetname="User_Financial_Input";
                else if(Tokens[i].equals("SOOCK")||Tokens[i].equals("EOOCK"))
                    sheetname="CrossCheck";
                else if(Tokens[i].equals("SORQ")||Tokens[i].equals("EORQ"))
                    sheetname="Reporting_Qtr";
                else if(Tokens[i].equals("SORY")||Tokens[i].equals("EORY"))
                    sheetname="Reporting_Year";
                else if(Tokens[i].equals("SOOCQ")||Tokens[i].equals("EOOCQ")||Tokens[i].equals("SOFCQ")||Tokens[i].equals("EOFCQ"))
                    sheetname="Chart_Qtr";
                else if(Tokens[i].equals("SOOCY")||Tokens[i].equals("EOOCY")||Tokens[i].equals("SOFCY")||Tokens[i].equals("EOFCY"))
                    sheetname="Chart_Year";
                else if(Tokens[i].equals("SOS")||Tokens[i].equals("EOS"))
                    sheetname="Summary";
                ErrorModel errmodel = new ErrorModel();
                errmodel.setError_desc("Token " + Tokens[i] + " is not Present");
                errmodel.setSheet_name(sheetname);
                //errmodel.setRow(-1);
                errmodel.setError_level("Error");
                errorModelList.add(errmodel);
                
            }
        }
		return errorModelList;
    }

    /**
     * 
     * @param sheetName name of sheet
     * @return true if sheet name is correct else return false 
     */
    private boolean containSheetName(String sheetName) {
    String[] correctSheetList={"BasicInfo","User_Operation_Input","User_Financial_Input","Unit_Map","Operation_Standard","Financial_Standard","Reporting_Qtr","Reporting_Year","CrossCheck","Chart_Qtr","Chart_Year","Summary","MetaDataSheet"};
    int found=0;
    for(int i=0;i<correctSheetList.length;i++)
    {
     if(sheetName.equals(correctSheetList[i]))
     {
         found=1;
         break;
     }
    }
    if(found==1)
    return true;
    else 
    return false;
    }
}
   

