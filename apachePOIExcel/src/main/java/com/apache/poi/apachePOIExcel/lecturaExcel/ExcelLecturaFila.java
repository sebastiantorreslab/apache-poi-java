package com.apache.poi.apachePOIExcel.lecturaExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;
import java.util.Iterator;

public class ExcelLecturaFila {

    public static void main(String[] args) {

        File archivo = new File("data.xlsx");

        try {
            InputStream input = new FileInputStream(archivo);

            XSSFWorkbook libro = new XSSFWorkbook(input);

            XSSFSheet hoja = libro.getSheetAt(1); // o getSheetAt(index)

            Row fila = hoja.getRow(0);

            Iterator<Cell> columnas = fila.cellIterator();

            //cuando leemos un documento tenemos que tener en cuenta el tipo de dato que estamos manejando

            while(columnas.hasNext()){
               Cell celda = columnas.next();

               if(celda.getCellType() == CellType.STRING){
                   String valor = celda.getStringCellValue();
                   System.out.println(valor);
               }
               if(celda.getCellType() == CellType.NUMERIC){
                   double valor = celda.getNumericCellValue();
                   System.out.println(valor);

               }if(celda.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(celda)){
                   Date fecha = celda.getDateCellValue();
                    System.out.println(fecha);
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
