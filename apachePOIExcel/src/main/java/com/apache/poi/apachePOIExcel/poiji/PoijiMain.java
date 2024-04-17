package com.apache.poi.apachePOIExcel.poiji;

import java.io.File;

public class Poiji {

    public static void main(String[] args) {

        File archivo = new File("data.xlsx");

        List<Persona> personas = Poiji.fromExcel();
    }
}
