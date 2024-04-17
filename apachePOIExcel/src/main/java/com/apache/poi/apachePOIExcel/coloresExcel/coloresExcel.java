package com.apache.poi.apachePOIExcel.coloresExcel;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class coloresExcel {

    public static void main(String[] args) {

        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja = libro.createSheet("colores");
        XSSFRow fila = hoja.createRow(1);
        XSSFCell celda = fila.createCell(1);
        XSSFCellStyle estilos = libro.createCellStyle();

        /* Configuración de estilos*/

        estilos.setFillForegroundColor(IndexedColors.CORAL.getIndex());
        estilos.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        /* Configuración de celda*/

        celda.setCellValue("Color predeterminado");
        celda.setCellStyle(estilos);

        /* Configuración hoja*/

        hoja.autoSizeColumn(1);

        try{
            OutputStream output = new FileOutputStream("ColoresExcel.xlsx");
            libro.write(output);
            libro.close();
            output.close();

        } catch (Exception e){
            e.printStackTrace();
        }

    }
}
