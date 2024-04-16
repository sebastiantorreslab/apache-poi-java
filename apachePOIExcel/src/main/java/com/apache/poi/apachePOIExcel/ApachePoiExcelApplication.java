package com.apache.poi.apachePOIExcel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileOutputStream;
import java.io.OutputStream;

@SpringBootApplication
public class ApachePoiExcelApplication {

    public static void main(String[] args) {

		/*

		 como crear un libro de excel VERSIÓN 2007 EN ADELANTE XLSX
		tra forma de crear un libro de excel VERSIÓN 1997 A 2003 EN ADELANTE XLS
		Workbook libro2 = new HSSFWorkbook();
		 No se puede trabajar con los dós al mismo tiempo, tienes que elegir el uno o el otro, usaremos XLSX
		*
		* */
        Workbook libro = new XSSFWorkbook();

        Sheet hoja1 = libro.createSheet("Personas");
        Sheet hoja2 = libro.createSheet("Contactos");
        Sheet hoja3 = libro.createSheet("Direcciones");

        try {
            OutputStream output = new FileOutputStream("ArchivoExcel.xlsx");
            libro.write(output);
        } catch (Exception e) {
            e.printStackTrace(); // todo: replace with a robust login
        }


        SpringApplication.run(ApachePoiExcelApplication.class, args);
    }

}
