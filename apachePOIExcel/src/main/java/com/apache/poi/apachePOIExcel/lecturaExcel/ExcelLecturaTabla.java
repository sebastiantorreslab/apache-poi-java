package com.apache.poi.apachePOIExcel.lecturaExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

public class ExcelLecturaTabla {

    public static void main(String[] args) {

        File archivo = new File("data.xlsx");

        try {
            InputStream input = new FileInputStream(archivo);

            XSSFWorkbook libro = new XSSFWorkbook(input);

            XSSFSheet hoja = libro.getSheetAt(2);

            Iterator<Row> filas = hoja.rowIterator();
            Iterator<Cell> columnas = null;

            Row filaActual = null;
            Cell columnaActual = null;

            while(filas.hasNext()){

                filaActual = filas.next();
                columnas = filaActual.cellIterator();

                while(columnas.hasNext()){

                    columnaActual = columnas.next();

                    if(columnaActual.getCellType() == CellType.STRING){
                        String valor = columnaActual.getStringCellValue();
                        System.out.println(valor);
                    }
                    if(columnaActual.getCellType() == CellType.NUMERIC){
                        double valor = columnaActual.getNumericCellValue();
                        System.out.println(valor);

                    }if(columnaActual.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(columnaActual)){
                        Date fecha = columnaActual.getDateCellValue();
                        SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
                        System.out.println(formato.format(fecha));
                    }

                }

                input.close();
                libro.close();

            }









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
