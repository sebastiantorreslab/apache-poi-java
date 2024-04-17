package com.apache.poi.apachePOIExcel.rangos;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class Rangos {

    public static void main(String[] args) {

        XSSFWorkbook libro1 = new XSSFWorkbook();
        XSSFSheet hoja = libro1.createSheet("rangos");
        XSSFRow fila1 = hoja.createRow(1);
        XSSFCell celda1 = fila1.createCell(1);
        XSSFCellStyle estilos1 = libro1.createCellStyle();

        CellRangeAddress rango = new CellRangeAddress(1,5,1,5);

        /* Configuración de estilos*/





        /* Configuración de celda*/




        /* Configuración hoja*/

        hoja.autoSizeColumn(1);
        hoja.addMergedRegion(rango);

        try{
            OutputStream output1 = new FileOutputStream("RangosExcel.xlsx");
            libro1.write(output1);
            libro1.close();
            output1.close();

        } catch (Exception e){
            e.printStackTrace();
        }



    }
}
