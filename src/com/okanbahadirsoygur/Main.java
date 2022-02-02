package com.okanbahadirsoygur;

import org.apache.poi.hssf.model.Workbook;
import org.apache.poi.hssf.record.formula.functions.Cell;
import org.apache.poi.hssf.record.formula.functions.Row;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;

public class Main {

    public static void main(String[] args) {
	// write your code here

        Scanner userInput = new Scanner(System.in);
        String komut = "0";
        while (true && !komut.equals( "3")){

            System.out.println("***    Java Excel Transactions Application ***");
            System.out.println("* (1)        EXCEL DOSYASINI KONTROL ET    ***");
            System.out.println("* (2)           EXCEL DOSYASINI OKU        ***");
            System.out.println("* (3)                 ÇIKIŞ YAP            ***");
            System.out.println("**********************************************");

            komut = userInput.nextLine();

            menuKontrol(komut,userInput);



        }//while

    }//main


    public static void konsolTemizle(){

        System.out.print("\033[H\033[2J");
        System.out.flush();

    }

    public static void menuKontrol(String komut, Scanner userInput){

        switch (komut){

            case "1":
                konsolTemizle();
                excelKontrol(userInput);

            break;

            case "2":
                excelDosyasiniOku(userInput);
            break;

            case "3":
            cikis();
            break;

        }//swich

    }


    public static void excelKontrol(Scanner userInput){
        File f = new File("data.xls");

        if(f.exists() && !f.isDirectory()) {
            System.out.println("data.xls Dosyası Kontrol Ediliyor...");
            System.out.println("Dosya Bulundu.");
            System.out.println("Başarılı.");
            System.out.println("Anamenü'ye dönmek için ENTER tuşuna basın...");
            userInput.nextLine();


        }else{


            System.out.println("data.xls Dosyası Kontrol Ediliyor...");
            System.out.println("Dosya Bulunamadı!");
            System.out.println("Taslak bir data.xls dosyası yaratılacaktır.");
            excelDosyasiYarat();
            System.out.println("Anamenü'ye dönmek için ENTER tuşuna basın...");

            userInput.nextLine();



        }

    }


    public static void excelDosyasiniOku(Scanner userInput){
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("data.xls"));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;


            int rows; // No of rows
            rows = sheet.getPhysicalNumberOfRows();

            int cols = 0; // No of columns
            int tmp = 0;

            // This trick ensures that we get the data properly even if it doesn't start from first few rows
            for(int i = 0; i < 10 || i < rows; i++) {
                row = sheet.getRow(i);
                if(row != null) {
                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                    if(tmp > cols) cols = tmp;
                }
            }

            for(int r = 0; r < rows; r++) {
                row = sheet.getRow(r);
                if(row != null) {
                    for(int c = 0; c < cols; c++) {
                        cell = row.getCell((short)c);
                        if(cell != null) {

                            System.out.println(cell+"");
                        }
                    }
                }
            }

        }catch (Exception e){

            System.out.println(e+"");

        }

        System.out.println("Anamenü'ye dönmek için ENTER tuşuna basın...");

        userInput.nextLine();
    }


    public static void excelDosyasiYarat(){
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Java-Excel-Transactions");



        HSSFRow row = sheet.createRow(0);

        HSSFCell cell = row.createCell((short)0);

        cell.setCellValue("Okan Bahadır Soygür");


        try {
            FileOutputStream out =
                    new FileOutputStream(new File("data.xls"));
            workbook.write(out);
            out.close();
            System.out.println("Excel Dosyası Başarıyla Yaratıldı.");
    }catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }


    public static void cikis(){

        System.out.println("Başarıyla Çıkış Yapıldı.");

    }



}
