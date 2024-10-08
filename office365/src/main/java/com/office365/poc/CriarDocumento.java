package com.office365.poc;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class CriarDocumento {

    public static void main(String[] args) {
        try {
            // Abrir o arquivo existente
            FileInputStream fis = new FileInputStream("C:\\Users\\vinic\\OneDrive\\Documentos\\arquivo.docx");
            XWPFDocument document = new XWPFDocument(fis);

            // Criar um novo parágrafo com texto
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun().setText("Este é o novo texto que estou adicionando ao arquivo Word.");

            // Salvar o arquivo atualizado
            FileOutputStream fos = new FileOutputStream("C:\\Users\\vinic\\OneDrive\\Documentos\\arquivo_atualizado.docx");
            document.write(fos);

            // Fechar streams
            fos.close();
            fis.close();

            System.out.println("Arquivo atualizado com sucesso!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
