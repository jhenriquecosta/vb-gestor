package com.bcptecnologias.migracao.gestortributario;

import java.awt.BorderLayout;

import javax.swing.JButton;
import javax.swing.JFrame;

public class Formulario extends JFrame{
	private JButton botao=new JButton();
	public static void main (String [] args){
		new Formulario().setVisible(true);
	}
	public void jbInit(){
		botao.setText("nome");
		this.getContentPane().setLayout(new BorderLayout());
		this.getContentPane().add(botao,BorderLayout.NORTH);
	}

}
