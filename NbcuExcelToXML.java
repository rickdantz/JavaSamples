/*
   Author.: Derrick Dantzler
   Date...: 04/23/2018
   Purpose: To read an MS Excel file and convert it into XML structure for IBM Sterling to perform
            further data mapping / validation.
   Notes..: For XML tags to be more meaningful the Excel should have Headers in the first row. 
   
   Structure of XML document should be: 
	   <ExcelRoot>
          <ExcelData>
             <CellLabel>SomeValue</CellLabel>
             <CellLabel>SomeValue</CellLabel>
             <CellLabel>SomeValue</CellLabel>
             <CellLabel>SomeValue</CellLabel>
          </ExcelData>
       </ExcelRoot> 
	
*/

package com.nbcu.sterling.extensions;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.*;

public class NBCUExcel2007ToXml {
	

 // This Method Reads the Excel file and creates and ArrayList of it's rows. 
 public static void readExcel(){

    FileInputStream inputStream = null;
    Workbook workbook = null;
    String excelFilePath = "C:\\XMLData\\SampleSheet.xlsx";
    
    // Begin File I/O
	try {
		inputStream = new FileInputStream(new File(excelFilePath));
	} catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
      
	// Open Excel Workbook and Trap for Errors
	try {
		workbook = new XSSFWorkbook(inputStream);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	
	 Sheet firstSheet = workbook.getSheetAt(0);
     Iterator<Row> iterator = firstSheet.iterator();
     
     // "data" is a multi-dimensional array the contains the <rows> and <the cell for each row>
     ArrayList<ArrayList<String>> data = new ArrayList<ArrayList<String>>();
     
     while (iterator.hasNext()) {
         Row nextRow = iterator.next();
         Iterator<Cell> cellIterator = nextRow.cellIterator();
          
         //This is an array for the cell values that get added to "data" array at the end of each loop
         ArrayList<String> cellData = new ArrayList<String>();
         
         // For every row check each cell that has data values
         while (cellIterator.hasNext()) {
             Cell cell = cellIterator.next();
             
             //Identify Each Cell Type and handle accordingly 
             switch (cell.getCellTypeEnum()){
                 case STRING: {
              	   //Cell has String data
              	   System.out.print(cell.getStringCellValue()); 
              	   cellData.add(cell.getStringCellValue());
              	   break; 
                 }
			       case BOOLEAN: {
			    	   //Cell has Boolean data
	            	   System.out.print(cell.getBooleanCellValue()); 
	            	   //Convert Boolean values to text 'true' or 'false'
	            	   if (cell.getBooleanCellValue() == true){
	            		   cellData.add("true");
	            		   cellData.clear();
	            	   }else {
	            		   cellData.add("false");
	            	   }
			    	   break;
			       }
			       case NUMERIC: {
			    	   //Cell has Numeric data
	            	   System.out.print(cell.getNumericCellValue()); 
	            	   cellData.add(cell.getNumericCellValue () + "");

			    	   break;
			       }
	               case BLANK: {
	            	   //Cell has Blank data
	            	   break;
	               }
			       case ERROR: {
			    	   //Cell has Error data
			    	   break;
			       }
			       case FORMULA: {
			    	   //Cell has formula data
			    	   break;
			       }
			       case _NONE: {
			    	   //Not sure what this is
			    	   break;
			       }
			       default: {
			    	   //Cell has data not defined 
			    	   break;
			       }
				        
             } // End of Switch Statement
            
             // These print() is just for debugging to see the results in the console
             System.out.print(" - ");
             
         } // End of While # 2
         
         System.out.println();
         
         //This is where the cell values get added to the multi-dimensional array
         data.add(cellData);
         
     } // End of While #1 
     
    try {
		workbook.close();
	    inputStream.close();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
 
    // Add code here to create XML
    try {
		@SuppressWarnings("unused")
		String isSuccess = buildXML(data);
	} catch (Throwable e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
     
     
 } // End Read Excel
 
 // This Method Receives and ArrayList and Creates and XML document
 public static String buildXML(ArrayList<ArrayList<String>> data) throws Throwable{
	 
	 int numOfRows = data.size();
	 
	 // Create new XML Object
	 DocumentBuilderFactory factory =DocumentBuilderFactory.newInstance();
	 DocumentBuilder builder = factory.newDocumentBuilder();
	 Document document = builder.newDocument();
	 
	 // Set Root Element of the XML Doc
	 Element rootElement = document.createElement("ExcelRoot");
	 document.appendChild(rootElement); 

	 // Create the Repeating Data Node > This will repeat for each row in the sheet
	 for (int i = 1; i < numOfRows; i++ ){
		 
		 Element dataElement = document.createElement("ExcelData");
		 rootElement.appendChild(dataElement);	 
		 
		 // Create the child Nodes that represent each cell on the sheet
		 int subIndex = data.get(i).size();
		 for (int index = 0; index < subIndex;){
			 
			  // Get the value of Column Header at this index and use as XML tag
			  String headerAsElementTag = data.get(0).get(index).toString();
			  Element rowElement = document.createElement(headerAsElementTag);
			  
			  // Now assign the cell data to this XML element
              dataElement.appendChild(rowElement);
              rowElement.appendChild(document.createTextNode(data.get(i).get(index)));
              index++;
		 }
		 
	 }

	   // This section outputs the XML document to a file 
       TransformerFactory tFactory = TransformerFactory.newInstance();
       Transformer transformer = tFactory.newTransformer();
       //Add indentation to output
       transformer.setOutputProperty
       (OutputKeys.INDENT, "yes");
       transformer.setOutputProperty(
             "{http://xml.apache.org/xslt}indent-amount", "2");

       DOMSource source = new DOMSource(document);
       StreamResult result = new StreamResult(new File("siprimarydoc.xml"));
       transformer.transform(source, result);
	 
       return "Success";
 }
 
 public static void main(String[] args) throws IOException {
	 
	 readExcel();
   } 
   
 
}
