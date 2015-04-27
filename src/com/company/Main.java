package com.company;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.XSLFSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;


import java.io.FileInputStream;

public class Main {

    public static void main(String[] args) throws Throwable {
        WordExtractor extractor = new WordExtractor(new FileInputStream("/home/mariusz/prez/dsds.doc"));
        System.out.println(extractor.getText());

        System.out.println("\n\n-----------------------------------------------------\n\n");

        XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor( new XWPFDocument(new FileInputStream("/home/mariusz/prez/dsds.docx")));
        System.out.println(xwpfWordExtractor.getText());

        System.out.println("\n\n-----------------------------------------------------\n\n");

        ExcelExtractor excelExtractor = new ExcelExtractor(new POIFSFileSystem(new FileInputStream("/home/mariusz/prez/ark1.xls")));
        System.out.println(excelExtractor.getText());

        System.out.println("\n\n-----------------------------------------------------\n\n");

        XSSFExcelExtractor xssfExcelExtractor = new XSSFExcelExtractor(new XSSFWorkbook(new FileInputStream("/home/mariusz/prez/ark1.xlsx")));
        System.out.println(xssfExcelExtractor.getText());

        System.out.println("\n\n-----------------------------------------------------\n\n");

        PowerPointExtractor powerPointExtractor = new PowerPointExtractor(new FileInputStream("/home/mariusz/prez/prez.ppt"));
        System.out.println(powerPointExtractor.getText());

        System.out.println("\n\n-----------------------------------------------------\n\n");

        XSLFPowerPointExtractor xslfPowerPointExtractor = new XSLFPowerPointExtractor(new XSLFSlideShow("/home/mariusz/prez/prez.pptx"));
        System.out.println(xslfPowerPointExtractor.getText());
    }
}
