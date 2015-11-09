/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/*
switch(cell_F.getCellType())
                        {
                            case Cell.CELL_TYPE_FORMULA:
                                reporting_Qtr_Formula_Cell_Formula=cell_F.getCellFormula();
                                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Qtr_Formula_Cell_Formula.charAt(19)=='G'))
                                  { 
                                      if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Qtr_Formula_Cell_Formula.substring(20, reporting_Qtr_Formula_Cell_Formula.length()))))
                                      {
                                        ErrorModel errorModel=new ErrorModel();
                                        CellReference cellRef=new CellReference(cell_F);
                                        errorModel.setCell_ref(cellRef.formatAsString());                             
                                        errorModel.setSheet_name("Reporting_Qtr");
                                        errorModel.setError_desc("Formula is Incorrect");
                                        errorModel.setError_level("Error");
                                        errorModelList.add(errorModel);    
                                      }
                                  }
                                 //else throw an error
                                  else
                                  {
                                  ErrorModel errorModel=new ErrorModel();
                                  CellReference cellRef=new CellReference(cell_F);
                                  errorModel.setCell_ref(cellRef.formatAsString());                             
                                  errorModel.setSheet_name("Reporting_Qtr");
                                  errorModel.setError_desc("Formula is Incorrect");
                                  errorModel.setError_level("Error");
                                  errorModelList.add(errorModel);                                   
                                  }         
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                break;
                            default:
                                   ErrorModel errorModel=new ErrorModel();
                                   CellReference cellRef=new CellReference(cell_F);
                                   errorModel.setCell_ref(cellRef.formatAsString());                             
                                   errorModel.setSheet_name("Reporting_Qtr");
                                   errorModel.setError_desc("This cell does not contain formula");
                                   errorModel.setError_level("Error");
                                   errorModelList.add(errorModel);
                                   break;
                        }
                switch(cell_G.getCellType())
                        {
                            case Cell.CELL_TYPE_FORMULA:
                                reporting_Qtr_Formula_Cell_Formula=cell_G.getCellFormula();
                                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Qtr_Formula_Cell_Formula.charAt(19)=='H'))
                                  { 
                                    if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Qtr_Formula_Cell_Formula.substring(20, reporting_Qtr_Formula_Cell_Formula.length()))))
                                    {
                                     ErrorModel errorModel=new ErrorModel();
                                    CellReference cellRef=new CellReference(cell_G);
                                    errorModel.setCell_ref(cellRef.formatAsString());                             
                                    errorModel.setSheet_name("Reporting_Qtr");
                                    errorModel.setError_desc("Formula is Incorrect");
                                    errorModel.setError_level("Error");
                                    errorModelList.add(errorModel);      
                                    }
                                  }
                                 //else throw an error
                                  else
                                  {
                                  ErrorModel errorModel=new ErrorModel();
                                  CellReference cellRef=new CellReference(cell_G);
                                  errorModel.setCell_ref(cellRef.formatAsString());                             
                                  errorModel.setSheet_name("Reporting_Qtr");
                                  errorModel.setError_desc("Formula is Incorrect");
                                  errorModel.setError_level("Error");
                                  errorModelList.add(errorModel);                                   
                                  }         
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                break;
                            default:
                                   ErrorModel errorModel=new ErrorModel();
                                   CellReference cellRef=new CellReference(cell_G);
                                   errorModel.setCell_ref(cellRef.formatAsString());                             
                                   errorModel.setSheet_name("Reporting_Qtr");
                                   errorModel.setError_desc("This cell does not contain formula");
                                   errorModel.setError_level("Error");
                                   errorModelList.add(errorModel);
                                   break;
                        }
                switch(cell_H.getCellType())
                        {
                            case Cell.CELL_TYPE_FORMULA:
                                reporting_Qtr_Formula_Cell_Formula=cell_H.getCellFormula();
                                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Qtr_Formula_Cell_Formula.charAt(19)=='I'))
                                  { 
                                      if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Qtr_Formula_Cell_Formula.substring(20, reporting_Qtr_Formula_Cell_Formula.length()))))
                                      {
                                       ErrorModel errorModel=new ErrorModel();
                                       CellReference cellRef=new CellReference(cell_H);
                                       errorModel.setCell_ref(cellRef.formatAsString());                             
                                       errorModel.setSheet_name("Reporting_Qtr");
                                       errorModel.setError_desc("Formula is Incorrect");
                                       errorModel.setError_level("Error");
                                       errorModelList.add(errorModel);   
                                      }
                                          
                                  } 
                                 //else throw an error
                                  else
                                  {
                                  ErrorModel errorModel=new ErrorModel();
                                  CellReference cellRef=new CellReference(cell_H);
                                  errorModel.setCell_ref(cellRef.formatAsString());                             
                                  errorModel.setSheet_name("Reporting_Qtr");
                                  errorModel.setError_desc("Formula is Incorrect");
                                  errorModel.setError_level("Error");
                                  errorModelList.add(errorModel);                                   
                                  }         
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                break;
                            default:
                                   ErrorModel errorModel=new ErrorModel();
                                   CellReference cellRef=new CellReference(cell_H);
                                   errorModel.setCell_ref(cellRef.formatAsString());                             
                                   errorModel.setSheet_name("Reporting_Qtr");
                                   errorModel.setError_desc("This cell does not contain formula");
                                   errorModel.setError_level("Error");
                                   errorModelList.add(errorModel);
                                break;
                        }
                switch(cell_I.getCellType())
                        {
                            case Cell.CELL_TYPE_FORMULA:
                                reporting_Qtr_Formula_Cell_Formula=cell_I.getCellFormula();
                                if((cell_B.getCellType()==Cell.CELL_TYPE_FORMULA)&&(reporting_Qtr_Formula_Cell_Formula.charAt(19)=='J'))
                                  { 
                                      if(!(cell_B.getCellFormula().substring(20, cell_B.getCellFormula().length()).equals(reporting_Qtr_Formula_Cell_Formula.substring(20, reporting_Qtr_Formula_Cell_Formula.length()))))
                                      {
                                        ErrorModel errorModel=new ErrorModel();
                                        CellReference cellRef=new CellReference(cell_I);
                                        errorModel.setCell_ref(cellRef.formatAsString());                             
                                        errorModel.setSheet_name("Reporting_Qtr");
                                        errorModel.setError_desc("Formula is Incorrect");
                                        errorModel.setError_level("Error");
                                        errorModelList.add(errorModel);    
                                      }
                                  } 
                                 //else throw an error
                                  else
                                  {
                                  ErrorModel errorModel=new ErrorModel();
                                  CellReference cellRef=new CellReference(cell_I);
                                  errorModel.setCell_ref(cellRef.formatAsString());                             
                                  errorModel.setSheet_name("Reporting_Qtr");
                                  errorModel.setError_desc("Formula is Incorrect");
                                  errorModel.setError_level("Error");
                                  errorModelList.add(errorModel);                                  
                                  }         
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                break;
                            default:
                                   ErrorModel errorModel=new ErrorModel();
                                   CellReference cellRef=new CellReference(cell_I);
                                   errorModel.setCell_ref(cellRef.formatAsString());                             
                                   errorModel.setSheet_name("Reporting_Qtr");
                                   errorModel.setError_desc("This cell does not contain formula");
                                   errorModel.setError_level("Error");
                                   errorModelList.add(errorModel);
                                break;
                        }
                
*/