/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package in.expertsoftware.colorcheck;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dharam
 */
public class FormatvarificationErrorList {
    public void dumpFormatErrorToExcelFile(ArrayList<ErrorModel> get_errormodelList)throws FileNotFoundException, IOException {
    /***
        Dump the error list into Excel file.***/
       XSSFWorkbook ErrorWorkbook = new XSSFWorkbook(); 
       XSSFSheet ErrorSheet;
       for(int i=0;i<get_errormodelList.size();i++) 
       {                     
           int index = ErrorWorkbook.getSheetIndex(get_errormodelList.get(i).sheet_name);
		if(index==-1)
                {
                    ErrorSheet=ErrorWorkbook.createSheet(get_errormodelList.get(i).sheet_name);  
                    XSSFRow totalError = ErrorSheet.createRow(0);
                    XSSFRow totalWarning = ErrorSheet.createRow(1);
                    CreaateHeaderOfErrorList(ErrorWorkbook,totalError.createCell(0),"Total Errors");
                    CreaateHeaderOfErrorList(ErrorWorkbook,totalWarning.createCell(0),"Total Warnings");
                    if((get_errormodelList.get(i).error_level).equals("Warning"))
                    { totalWarning.createCell(1).setCellValue(1);  }
                    else
                        totalWarning.createCell(1).setCellValue(0); 
                   if((get_errormodelList.get(i).error_level).equals("Error"))
                   {    totalError.createCell(1).setCellValue(1);  }
                   else
                    totalError.createCell(1).setCellValue(0);

                    ErrorSheet.createRow(2);
                    XSSFRow headerrow = ErrorSheet.createRow(3);
                    //create header of every sheet
                    Cell Header_referenceCell =headerrow.createCell(0);
                    CreaateHeaderOfErrorList(ErrorWorkbook,Header_referenceCell,"Cell Ref");
                    Cell Header_sheetname =headerrow.createCell(1);
                    CreaateHeaderOfErrorList(ErrorWorkbook,Header_sheetname,"Sheet Name"); 
                    Cell Header_ErrorDesc =headerrow.createCell(2);
                    CreaateHeaderOfErrorList(ErrorWorkbook,Header_ErrorDesc,"Error Description"); 
                    Cell Header_ErrorLevel =headerrow.createCell(3);
                    CreaateHeaderOfErrorList(ErrorWorkbook,Header_ErrorLevel,"Error Level");                                           
                    XSSFRow row = ErrorSheet.createRow(4);
                    row = ErrorSheet.createRow(5);
                    
                    CreaateStyleOfErrorList(ErrorWorkbook,row,get_errormodelList.get(i).cell_ref,get_errormodelList.get(i).sheet_name,get_errormodelList.get(i).error_desc,get_errormodelList.get(i).error_level);  
                    ErrorSheet.autoSizeColumn(0);
                    ErrorSheet.autoSizeColumn(1);
                    ErrorSheet.autoSizeColumn(2);
                    ErrorSheet.autoSizeColumn(3);
               }
                else
                { 
                 printinfo(ErrorWorkbook,get_errormodelList.get(i).cell_ref,get_errormodelList.get(i).sheet_name,get_errormodelList.get(i).error_desc,get_errormodelList.get(i).error_level);       
                }                    
       } 
        setColorInfoMetaData(ErrorWorkbook);      
        try (FileOutputStream ErrorOutputStream = new FileOutputStream("C:/Users/Dharam/Desktop/DIMT_NEW_CODE/ErrorSheet.xlsx")) {
            ErrorWorkbook.write(ErrorOutputStream);        
        }
    }
    private static void printinfo(XSSFWorkbook ErrorWorkbook, String cell_ref, String sheet_name, String error_desc, String error_level) 
    {
    XSSFSheet ErrorSheet= ErrorWorkbook.getSheet(sheet_name);
    if(error_level.equals("Error"))
    {
        ErrorSheet.getRow(0).getCell(1).setCellValue(1+ErrorSheet.getRow(0).getCell(1).getNumericCellValue());
    }
    else if(error_level.equals("Warning"))
    {
        ErrorSheet.getRow(1).getCell(1).setCellValue(1+ErrorSheet.getRow(1).getCell(1).getStringCellValue());
    }
    XSSFRow row = ErrorSheet.createRow(ErrorSheet.getPhysicalNumberOfRows());
    CreaateStyleOfErrorList(ErrorWorkbook,row,cell_ref,sheet_name,error_desc,error_level);
    
    }
    private static void CreaateHeaderOfErrorList(XSSFWorkbook ErrorWorkbook,Cell column,String text) {
                    XSSFCellStyle headerStyleOfreference = ErrorWorkbook.createCellStyle();
                    headerStyleOfreference.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                    headerStyleOfreference.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    headerStyleOfreference.setFillForegroundColor(new XSSFColor(new java.awt.Color(217, 217, 217)));
                    headerStyleOfreference.setBorderBottom((short) 1);
                    headerStyleOfreference.setBorderTop((short) 1);
                    headerStyleOfreference.setBorderLeft((short) 1);
                    headerStyleOfreference.setBorderRight((short) 1);
                    
                    //create font
                    XSSFFont fontOfCellFirst = ErrorWorkbook.createFont();
                    fontOfCellFirst.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                    fontOfCellFirst.setFontHeightInPoints((short) 12);
                    fontOfCellFirst.setFontName("Calibri");
                    fontOfCellFirst.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
                    headerStyleOfreference.setFont(fontOfCellFirst);
                    column.setCellValue(text);
                    column.setCellStyle(headerStyleOfreference);
        }
    
    private static void CreaateStyleOfErrorList(XSSFWorkbook ErrorWorkbook, XSSFRow row, String cell_ref, String sheet_name, String error_desc, String error_level) {
                    XSSFCellStyle StyleOfCell = ErrorWorkbook.createCellStyle();
                    StyleOfCell.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                    StyleOfCell.setFillPattern(FillPatternType.SOLID_FOREGROUND);                    
                    
                    if(error_level.equalsIgnoreCase("Warning"))
                    {StyleOfCell.setFillForegroundColor(new XSSFColor(new java.awt.Color(155, 194, 230)));}                    
                    else
                    {StyleOfCell.setFillForegroundColor(new XSSFColor(new java.awt.Color(225, 171, 171)));}    
                    StyleOfCell.setBorderLeft((short) 1);
                    StyleOfCell.setBorderRight((short) 1);
                    StyleOfCell.setBorderTop((short) 1);
                    StyleOfCell.setBorderBottom((short) 1);
                    StyleOfCell.setWrapText(true);
                   
                    //create font
                    XSSFFont fontOfCell = ErrorWorkbook.createFont();
                    fontOfCell.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                    fontOfCell.setFontHeightInPoints((short) 10);
                    fontOfCell.setFontName("Calibri");
                    fontOfCell.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
                    StyleOfCell.setFont(fontOfCell);
                    Cell Rowcell_0=row.createCell(0);                                                               
                    Cell Rowcell_1=row.createCell(1);                   
                    Cell Rowcell_2=row.createCell(2);                    
                    Cell Rowcell_3=row.createCell(3);                                        
                    Rowcell_0.setCellValue(cell_ref);
                    Rowcell_1.setCellValue(sheet_name);
                    Rowcell_2.setCellValue(error_desc);
                    Rowcell_3.setCellValue(error_level);                        
                    Rowcell_0.setCellStyle(StyleOfCell); 
                    Rowcell_1.setCellStyle(StyleOfCell); 
                    Rowcell_2.setCellStyle(StyleOfCell); 
                    Rowcell_3.setCellStyle(StyleOfCell); 
            }    

    private static void setColorInfoMetaData(XSSFWorkbook ErrorWorkbook) {
        //Set Colour used information on first sheet.
       XSSFSheet setInfoSheet=ErrorWorkbook.getSheetAt(0);
       XSSFRow colourInfoRow; 
       XSSFRow errorColourRow;
       XSSFRow warningColourRow;
       if(setInfoSheet.getPhysicalNumberOfRows()>5)    
       {colourInfoRow=setInfoSheet.getRow(5);}
       else
       {colourInfoRow=setInfoSheet.createRow(5);}          
       Cell colorInfoCell=colourInfoRow.createCell(6);
       Cell RGBCell=colourInfoRow.createCell(7);
       CreaateHeaderOfErrorList(ErrorWorkbook,colorInfoCell,"Used Color");
       CreaateHeaderOfErrorList(ErrorWorkbook,RGBCell,"RGB Value");
       setInfoSheet.autoSizeColumn(6);
       setInfoSheet.autoSizeColumn(7); 
       if(setInfoSheet.getPhysicalNumberOfRows()>6)
       {errorColourRow=setInfoSheet.getRow(6); }
       else {errorColourRow=setInfoSheet.createRow(6);}
       if(setInfoSheet.getPhysicalNumberOfRows()>7)
       {warningColourRow=setInfoSheet.getRow(7);}
       else{warningColourRow=setInfoSheet.createRow(7);}       
       //error color style
       XSSFCellStyle errorStyle = ErrorWorkbook.createCellStyle();       
       errorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
       errorStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(225, 171, 171)));
       errorColourRow.createCell(6).setCellStyle(errorStyle);
       errorColourRow.getCell(6).setCellValue("Error");
       errorColourRow.createCell(7).setCellValue("225, 171, 171");
       //warning color style
       XSSFCellStyle warningStyle = ErrorWorkbook.createCellStyle();       
       warningStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
       warningStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(155, 194, 230)));
       warningColourRow.createCell(6).setCellStyle(warningStyle);
       warningColourRow.getCell(6).setCellValue("Warning");
       warningColourRow.createCell(7).setCellValue("155, 194, 230");
       }
}
