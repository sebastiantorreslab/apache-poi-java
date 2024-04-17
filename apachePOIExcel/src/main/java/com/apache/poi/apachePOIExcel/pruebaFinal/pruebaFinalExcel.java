package com.apache.poi.apachePOIExcel.pruebaFinal;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class pruebaFinalExcel {

    public static void main(String[] args) {

        List<Cliente> listado = getList();
        Field[] campos = Cliente.class.getDeclaredFields();

        // patrón builder



        XSSFWorkbook libro = new XSSFWorkbook();

        XSSFSheet hoja = libro.createSheet("Clientes");

        XSSFCellStyle estiloTitulo = new GeneradorEstilos.Builder().setColorDefecto(IndexedColors.DARK_BLUE.getIndex())
                .setTipoPatron(FillPatternType.SOLID_FOREGROUND)
                .setAlineacionHorizontal(HorizontalAlignment.CENTER)
                .build(libro);


        XSSFRow fila = null;
        XSSFCell celda = null;

        for (int i = 0; i < listado.size(); i++) {
            if (i == 0) {
                fila = hoja.createRow(0);

                for (int j = 0; j < campos.length; j++) {
                    celda = fila.createCell(j);
                    celda.setCellValue(campos[j].getName());
                    celda.setCellStyle(estiloTitulo);
                }

            }

                Cliente cliente = listado.get(i);
                List<Object> atributos = cliente.obtenerAtributos();

                fila = hoja.createRow(i+1);
                for (int a = 0; a < atributos.size(); a++){

                    celda = fila.createCell(a);

                    if(atributos.get(a) instanceof Long){
                        celda.setCellValue((Long) atributos.get(a));
                    }
                    if(atributos.get(a) instanceof String){
                        celda.setCellValue((String) atributos.get(a));
                    }
                    if(atributos.get(a) instanceof LocalDate){
                        celda.setCellValue((LocalDate) atributos.get(a));
                    }

                    hoja.autoSizeColumn(a);



            }

        }

        try {

            OutputStream output = new FileOutputStream("pruebaFinal.xlsx");
            libro.write(output);

            libro.close();
            output.close();
        } catch (Exception e) {

            e.printStackTrace();
            throw new RuntimeException("Error creando el documento");


        }


    }


    public static List<Cliente> getList() {
        List<Cliente> listaClientes = new ArrayList<>();
        listaClientes.add(new Cliente(1L, "Sebastian", "Perez", "123456", "jj@email.com", LocalDate.of(1998, 11, 14)));
        listaClientes.add(new Cliente(2L, "John", "Dow", "123456", "mm@email.com", LocalDate.of(1995, 11, 24)));
        listaClientes.add(new Cliente(3L, "Robert", "Pires", "123456", "rr@email.com", LocalDate.of(2001, 1, 3)));
        return listaClientes;
    }
}
