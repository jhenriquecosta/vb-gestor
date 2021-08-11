package com.bcptecnologias.migracao.gestortributario.fachada;

import java.io.File;

public class ControladorD2Ti {
	private static File diretorio(){
		ClassLoader classLoader = ControladorD2Ti.class.getClassLoader();
		System.out.println(classLoader.getResource(""));
		return null;
	}
	public static void main(String[] args) {
		diretorio();
	}
}