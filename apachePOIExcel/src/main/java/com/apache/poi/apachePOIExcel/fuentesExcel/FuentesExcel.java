package com.apache.poi.apachePOIExcel.fuentesExcel;

import org.apache.poi.ss.formula.functions.Index;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class FuentesExcel {

    public static void main(String[] args) {

        XSSFWorkbook libro1 = new XSSFWorkbook();
        XSSFSheet hoja = libro1.createSheet("fuentes");
        XSSFRow fila1 = hoja.createRow(1);
        XSSFCell celda1 = fila1.createCell(1);
        XSSFCellStyle estilos1 = libro1.createCellStyle();

        /* Configuración de estilos*/

        XSSFFont fuente = libro1.createFont();
        fuente.setFontName("Franklin Gothic Book");
        fuente.setBold(true);
        fuente.setItalic(true);
        fuente.setFontHeightInPoints((short ) 14);
        fuente.setColor(IndexedColors.GREEN.getIndex());
        fuente.setUnderline(FontUnderline.DOUBLE);

        estilos1.setFont(fuente);
        estilos1.setAlignment(HorizontalAlignment.CENTER);


        /* Configuración de celda*/

        celda1.setCellValue("Fuente excel");
        celda1.setCellStyle(estilos1);

        /* Configuración hoja*/

        hoja.autoSizeColumn(1);

        try{
            OutputStream output1 = new FileOutputStream("FuentesExcel.xlsx");
            libro1.write(output1);
            libro1.close();
            output1.close();

        } catch (Exception e){
            e.printStackTrace();
        }

    }
}
