package com.apache.poi.apachePOIExcel.lecturaExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class ExcelLecturaColumna {

    public static void main(String[] args) {

        File archivo = new File("data.xlsx");

        try {
            InputStream input = new FileInputStream(archivo);

            XSSFWorkbook libro = new XSSFWorkbook(input);

            XSSFSheet hoja = libro.getSheetAt(0); // o getSheetAt(index)

            //Row fila = hoja.getRow(1); apra traer una fila

            Iterator<Row> filas = hoja.rowIterator();


            while(filas.hasNext()){

                Cell columna = filas.next().getCell(0);

                if(columna != null){
                    System.out.println(columna.getStringCellValue());
                }

            }

            input.close();
            libro.close();


        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }
}
