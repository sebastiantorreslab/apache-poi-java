package com.apache.poi.apachePOIExcel.escribirExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class Ejercicio1 {
    public static void main(String[] args) {

        Workbook libroEjercicios = new XSSFWorkbook();
        Sheet hoja = libroEjercicios.createSheet("Ejercicio1");
        Row headers = hoja.createRow(2);
        Row registro1 = hoja.createRow(3);
        Row registro2 = hoja.createRow(4);

        Cell nombre = headers.createCell(1);
        Cell edad = headers.createCell(2);
        Cell ciudad = headers.createCell(3);


        nombre.setCellValue("Nombre");
        edad.setCellValue("Edad");
        ciudad.setCellValue("Ciudad");


        Cell insertName1 = registro1.createCell(1);
        Cell insertAge1 = registro1.createCell(2);
        Cell insertCity1 =  registro1.createCell(3);

        insertName1.setCellValue("Santiago");
        insertAge1.setCellValue(23);
        insertCity1 .setCellValue("Manizales");

        Cell insertName2 = registro2.createCell(1);
        Cell insertAge2 = registro2.createCell(2);
        Cell insertCity2 =  registro2.createCell(3);

        insertName2.setCellValue("Angie");
        insertAge2.setCellValue(22);
        insertCity2.setCellValue("Bogotá");

        try {
            OutputStream output = new FileOutputStream("EjerciciosExcel.xlsx");
            libroEjercicios.write(output);
            libroEjercicios.close();
            output.close();
        } catch (Exception e) {
            e.printStackTrace(); // todo: replace with a robust login
        }




    }
}
