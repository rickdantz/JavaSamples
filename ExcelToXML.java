/*
  Sample from javaworld.com
  https://www.javaworld.com/article/2076189/enterprise-java/book-excerpt--converting-xml-to-spreadsheet--and-vice-versa.html 
  
*/

package com.apress.excel;

import org.apache.poi.hssf.usermodel.*;
import org.w3c.dom.*;
import java.io.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
public class ExcelToXML {
   public void generateXML(File excelFile) {
      try { //Initializing the XML document
         DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
         DocumentBuilder builder = factory.newDocumentBuilder();
         Document document = builder.newDocument();
         Element rootElement = document.createElement("incmstmts");
         document.appendChild(rootElement);
            //Creating top-level elements
         Element stmtElement1 = document.createElement("stmt");
         rootElement.appendChild(stmtElement1);

         Element stmtElement2 = document.createElement("stmt");
         rootElement.appendChild(stmtElement2);
            //Adding first subelements
         Element year1 = document.createElement("year");
         stmtElement1.appendChild(year1);

         year1.appendChild(document.createTextNode("2005"));

         Element year2 = document.createElement("year");
         stmtElement2.appendChild(year2);
         year2.appendChild(document.createTextNode("2004"));
            //Creating an HSSFSpreadsheet object from an Excel file
         InputStream input = new FileInputStream(excelFile);
         HSSFWorkbook workbook = new HSSFWorkbook(input);
         HSSFSheet spreadsheet = workbook.getSheetAt(0);

         for (int i = 1; i <= spreadsheet.getLastRowNum(); i++) {
            switch (i) {
         //Iterate over spreadsheet rows to create stmt element
         //subelements.
            case 1:
               HSSFRow row1 = spreadsheet.getRow(1);

         Element revenueElement1 = document.createElement("revenue");
            stmtElement1.appendChild(revenueElement1);

            revenueElement1.appendChild
            (document.createTextNode
            (row1.getCell((short) 1).
            getStringCellValue()));

         Element revenueElement2 = document.createElement("revenue");
         stmtElement2.appendChild(revenueElement2);

            revenueElement2.appendChild
            (document.createTextNode
            (row1.getCell((short) 2).
            getStringCellValue()));

         break;
         case 2:
            HSSFRow row2 = spreadsheet.getRow(2);

            Element costofrevenue1 = document.createElement("costofrevenue");
            stmtElement1.appendChild(costofrevenue1);
            costofrevenue1.appendChild
             (document.createTextNode
             (row2.getCell((short)1).
            getStringCellValue()));

         Element costofrevenue2 = document.createElement("costofrevenue");
            stmtElement2.appendChild(costofrevenue2);

            costofrevenue2.appendChild
            (document.createTextNode
            (row2.getCell((short) 2).
            getStringCellValue()));
            break;
         case 3:
            HSSFRow row3 = spreadsheet.getRow(3);

         Element researchdevelopment1 = document.createElement("researchdevelopment");
            stmtElement1.appendChild(researchdevelopment1);

            researchdevelopment1.appendChild
            (document.createTextNode
            (row3.getCell((short) 1).getStringCellValue()));

               Element researchdevelopment2 =document.createElement("researchdevelopment");
               stmtElement2.appendChild(researchdevelopment2);

               researchdevelopment2.appendChild
              (document.createTextNode
              (row3.getCell((short) 2).
               getStringCellValue()));
               break;
            case 4:
               HSSFRow row4 = spreadsheet.getRow(4);

               Element salesmarketing1 = document.createElement("salesmarketing");
               stmtElement1.appendChild(salesmarketing1);

            salesmarketing1.appendChild(document.createTextNode(row4.getCell((short) 1).getStringCellValue()));

               Element salesmarketing2 = document.createElement("salesmarketing");
               stmtElement2.appendChild(salesmarketing2);
               salesmarketing2.appendChild(document.createTextNode(row4.getCell((short) 2).getStringCellValue()));
               break;
            case 5:
               HSSFRow row5 = spreadsheet.getRow(5);

               Element generaladmin1 = document.createElement("generaladmin");
               stmtElement1.appendChild(generaladmin1);

               generaladmin1.appendChild(document.createTextNode(row5
                  .getCell((short) 1).getStringCellValue()));

               Element generaladmin2 = document.createElement("generaladmin");
               stmtElement2.appendChild(generaladmin2);

               generaladmin2.appendChild(document.createTextNode(row5
               .getCell((short) 2).getStringCellValue()));
               break;
            case 6:
               HSSFRow row6 = spreadsheet.getRow(6);

               Element totaloperexpenses1 = document.createElement("totaloperexpenses");
               stmtElement1.appendChild(totaloperexpenses1);

               totaloperexpenses1.appendChild(document.createTextNode(row6
                  .getCell((short) 1).getStringCellValue()));

               Element totaloperexpenses2 = document.createElement("totaloperexpenses");
               stmtElement2.appendChild(totaloperexpenses2);

               totaloperexpenses2.appendChild(document.createTextNode(row6
                  .getCell((short) 2).getStringCellValue()));
               break;
            case 7:
               HSSFRow row7 = spreadsheet.getRow(7);

            Element operincome1 = document.createElement("operincome");
               stmtElement1.appendChild(operincome1);

               operincome1.appendChild(document.createTextNode(row7
                  .getCell((short) 1).getStringCellValue()));

            Element operincome2 = document.createElement("operincome");
               stmtElement2.appendChild(operincome2);

               operincome2.appendChild
                (document.createTextNode
                (row7.getCell((short) 2).
               getStringCellValue()));
               break;
            case 8:
               HSSFRow row8 = spreadsheet.getRow(8);

            Element invincome1 = document.createElement("invincome");
               stmtElement1.appendChild(invincome1);

               invincome1.appendChild
                (document.createTextNode
                (row8.getCell((short) 1).
               getStringCellValue()));

            Element invincome2 = document.createElement("invincome");
            stmtElement2.appendChild(invincome2);

               invincome2.appendChild
                (document.createTextNode
                (row8.getCell((short) 2).
               getStringCellValue()));
               break;
            case 9:
               HSSFRow row9 = spreadsheet.getRow(9);

               Element incbeforetaxes1 = document.createElement("incbeforetaxes");
               stmtElement1.appendChild(incbeforetaxes1);

               incbeforetaxes1.appendChild
               (document.createTextNode
               (row9.getCell((short) 1).
               getStringCellValue()));

               Element incbeforetaxes2 =document.createElement("incbeforetaxes");
               stmtElement2.appendChild(incbeforetaxes2);

               incbeforetaxes2.appendChild
                (document.createTextNode
                (row9.getCell((short)2).
               getStringCellValue()));
               break;
            case 10:
               HSSFRow row10 = spreadsheet.getRow(10);

               Element taxes1 = document.createElement("taxes");
               stmtElement1.appendChild(taxes1);

               taxes1.appendChild(document.createTextNode(row10.getCell(
                   (short) 1).getStringCellValue()));

               Element taxes2 = document.createElement("taxes");
               stmtElement2.appendChild(taxes2);

               taxes2.appendChild(document.createTextNode(row10.getCell(
                   (short) 2).getStringCellValue()));
               break;

            case 11:
               HSSFRow row11 = spreadsheet.getRow(11);

            Element netincome1 = document.createElement("netincome");
               stmtElement1.appendChild(netincome1);

            netincome1.appendChild(document.createTextNode(row11
               .getCell((short) 1).getStringCellValue()));

            Element netincome2 = document.createElement("netincome");
               stmtElement2.appendChild(netincome2);

               netincome2.appendChild(document.createTextNode(row11
                  .getCell((short) 2).getStringCellValue()));
               break;
            default:
               break;
            }

         }

         TransformerFactory tFactory = TransformerFactory.newInstance();

         Transformer transformer = tFactory.newTransformer();
            //Add indentation to output
         transformer.setOutputProperty
         (OutputKeys.INDENT, "yes");
         transformer.setOutputProperty(
            "{http://xml.apache.org/xslt}indent-amount", "2");

         DOMSource source = new DOMSource(document);
         StreamResult result = new StreamResult(System.out);
         transformer.transform(source, result);
      } catch (IOException e) {
         System.out.println("IOException " + e.getMessage());
      } catch (ParserConfigurationException e) {
         System.out
            .println("ParserConfigurationException " + e.getMessage());
      } catch (TransformerConfigurationException e) {
         System.out.println("TransformerConfigurationException "+ e.getMessage());
      } catch (TransformerException e) {
         System.out.println("TransformerException " + e.getMessage());
      }
   }

   public static void main(String[] argv) {
      ExcelToXML excel = new ExcelToXML();
      File input = new File("IncomeStatements.xls");
      excel.generateXML(input);
   }
}
