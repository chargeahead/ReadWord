package org.example;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.List;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {
    public static void main(String[] args) throws IOException {
       String fileName = "/Users/pankajjain/Downloads/test.docx";

       try(XWPFDocument doc = new XWPFDocument(
               Files.newInputStream(Paths.get(fileName)))){
           //Read the word doc and search for a word
 /*
           XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
           String txt = extractor.getText();
           System.out.println("Word text: "+txt);
           if (txt.contains("Apache")){
               System.out.println("Search word Apache present");
           }

  */
           //Read just the paragraphs
           List<XWPFParagraph> paras = doc.getParagraphs();
           for(XWPFParagraph para: paras){
               System.out.println(para.getText());
           }

           //Iterate over the document element wise to just print the table
           Iterator<IBodyElement> iterator = doc.getBodyElementsIterator();
           while(iterator.hasNext()){
               IBodyElement element = iterator.next();
               if(element.getElementType().name().equalsIgnoreCase("TABLE")){
                   List<XWPFTable> tableList = element.getBody().getTables();
                   for(XWPFTable tbl: tableList){
                       for (int i = 0; i< tbl.getRows().size(); i++){
                           for (int j = 0; j < tbl.getRow(i).getTableCells().size(); j++){
                               System.out.println(tbl.getRow(i).getCell(j).getText());
                           }
                       }
                   }
               }
           }
       }
    }
}