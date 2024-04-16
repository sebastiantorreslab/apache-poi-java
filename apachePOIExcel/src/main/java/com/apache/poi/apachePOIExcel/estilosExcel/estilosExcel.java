package com.apache.poi.apachePOIExcel.estilosExcel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class estilosExcel {

    public static void main(String[] args) {

        // Cuando se necesite más configuraciones formatos y funcionalidades adicionlaes se debe crear el libro
        // de excel de esta forma. Pues así ya estará incluyendo los métodos propios de esta clase.

        XSSFWorkbook libro = new XSSFWorkbook();

        XSSFSheet hoja = libro.createSheet("Estilos");

        XSSFRow fila = hoja.createRow(1);

        //Siempre los estilos se dan a nivel de celda

        XSSFCell celda = fila.createCell(1);


        /*Configuración de estilos*/

        XSSFCellStyle estiloCelda = libro.createCellStyle();

        estiloCelda.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estiloCelda.setBorderRight(BorderStyle.DOTTED);
        estiloCelda.setBorderTop(BorderStyle.DOTTED);
        estiloCelda.setBorderBottom(BorderStyle.DOTTED);
        estiloCelda.setBorderLeft(BorderStyle.DOTTED);

        /*
        Configuración de celda
        * */

        celda.setCellValue("Estilos con apache POI");
        celda.setCellStyle(estiloCelda);
        hoja.autoSizeColumn(1); // es importante que la configuración de las hojas vaya despues de la configuración de las celdas


        try {

            OutputStream output = new FileOutputStream("EstilosExcel.xlsx");
            libro.write(output);
            libro.close();
            output.close();

        } catch (Exception e) {

            e.printStackTrace();

        }


    }
}
