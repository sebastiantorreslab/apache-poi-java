package com.apache.poi.apachePOIExcel.formatos;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.time.LocalDateTime;

public class Formatos {


    public static void main(String[] args) {

        XSSFWorkbook libro1 = new XSSFWorkbook();
        XSSFSheet hoja = libro1.createSheet("rangos");
        XSSFRow fila1 = hoja.createRow(1);
        XSSFCell celda1 = fila1.createCell(1);
        XSSFCellStyle estilos1 = libro1.createCellStyle();


        /* Configuración de estilos*/
        estilos1.setDataFormat(libro1.createDataFormat().getFormat("dd/MM/yyyy HH:mm:ss"));


        /* Configuración de celda*/
        celda1.setCellValue(LocalDateTime.now());
        celda1.setCellStyle(estilos1);


        /* Configuración hoja*/

        hoja.autoSizeColumn(1);

        try{
            OutputStream output1 = new FileOutputStream("FormatosExcel.xlsx");
            libro1.write(output1);
            libro1.close();
            output1.close();

        } catch (Exception e){
            e.printStackTrace();
        }



    }
}
