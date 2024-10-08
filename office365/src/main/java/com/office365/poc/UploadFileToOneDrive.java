package com.office365.poc;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

@SpringBootApplication
@RestController
public class UploadFileToOneDrive {

	public static void main(String[] args) {
		SpringApplication.run(PocApplication.class, args);
	}

	@GetMapping("/hello")
    public String atualizarDocumento(@RequestParam String idDocumento,@RequestParam String conteudo) throws IOException, InterruptedException {
		atualizarArquivo(conteudo);
		HttpClient httpClient = HttpClient.newHttpClient();

        // Enviar o arquivo atualizado de volta para o OneDrive
        HttpRequest request=null;
		try {
			request = HttpRequest.newBuilder()
			    .uri(URI.create("https://graph.microsoft.com/v1.0/me/drive/items/"+idDocumento+"/content"))
			    .header("Authorization", "Bearer " + "eyJ0eXAiOiJKV1QiLCJub25jZSI6Inp3cldjTlFWc3c4YXBNazhKeGRpbVR2aEVmSWw1eTA5U2FDWTA5Nkhxa28iLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1jN2wzSXo5M2c3dXdnTmVFbW13X1dZR1BrbyIsImtpZCI6Ik1jN2wzSXo5M2c3dXdnTmVFbW13X1dZR1BrbyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8xNzU2ZTU0MC1kZDUwLTRiYmQtODA1Yy02NDE3ZDI5MjgxYTIvIiwiaWF0IjoxNzI4MzI5NTM0LCJuYmYiOjE3MjgzMjk1MzQsImV4cCI6MTcyODQxNjIzNCwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFaUUFhLzhZQUFBQVhhM0lucDdlUjlpL3dRTkFsVlpkdmJZcXJ3UWNmTkxrS1kyS3k3dlpPcTFMRDdqdnFqZVA4WVZZZWlaSG1qWmRFRTh2emVEaTNWbVliRk95bEV0eDZ2SjhYTGJGd0dTSHU3TmVLN0hrSld6dGJKbFpISVY1bC8xNUNxNFBsRGlDSFQydU1jbXUyMStrbGhMVlBKK2tweVFjODNISkJOV1ZWV2VoYWRNUmJmRUdSM3daNngyZk1Ta3JVdTV1SDJseiIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiQ2FycmlvbiIsImdpdmVuX25hbWUiOiJWaW5pY2l1cyIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjE3Ny4xMTguMTU3LjExMCIsIm5hbWUiOiJWaW5pY2l1cyBDYXJyaW9uIiwib2lkIjoiMTdlNzI3MDAtYjNkZS00MjU2LWIxOTAtZGQ2OWU5NTI1ZjMyIiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAwQzIwQzVEQTAiLCJyaCI6IjAuQVZrQVFPVldGMURkdlV1QVhHUVgwcEtCb2dNQUFBQUFBQUFBd0FBQUFBQUFBQUJaQUtBLiIsInNjcCI6IkF1ZGl0TG9nc1F1ZXJ5LU9uZURyaXZlLlJlYWQuQWxsIEZpbGVzLlJlYWRXcml0ZSBGaWxlcy5SZWFkV3JpdGUuQWxsIG9wZW5pZCBwcm9maWxlIFNlcnZpY2VBY3Rpdml0eS1PbmVEcml2ZS5SZWFkLkFsbCBVc2VyLlJlYWQgZW1haWwgU2l0ZXMuUmVhZFdyaXRlLkFsbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6Ijd5WXpIX1k3ZmFSVTEwMkR5d3hybl9WTHRROE9MUTNTQ3RxTV8zQUFVQzAiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiU0EiLCJ0aWQiOiIxNzU2ZTU0MC1kZDUwLTRiYmQtODA1Yy02NDE3ZDI5MjgxYTIiLCJ1bmlxdWVfbmFtZSI6InZpbmljaXVzLmNhcnJpb25AZmFicmljYWRzLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6InZpbmljaXVzLmNhcnJpb25AZmFicmljYWRzLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6IlV3RUd6ek9sXzBXWm9ZbGszTE1jQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjNlZGFmNjYzLTM0MWUtNDQ3NS05Zjk0LTVjMzk4ZWY2YzA3MCIsIjc2OThhNzcyLTc4N2ItNGFjOC05MDFmLTYwZDZiMDhhZmZkMiIsIjJiNzQ1YmRmLTA4MDMtNGQ4MC1hYTY1LTgyMmM0NDkzZGFhYyIsIjQ0MzY3MTYzLWViYTEtNDRjMy05OGFmLWY1Nzg3ODc5Zjk2YSIsIjM4YTk2NDMxLTJiZGYtNGI0Yy04YjZlLTVkM2Q4YWJhYzFhNCIsIjE1OGMwNDdhLWM5MDctNDU1Ni1iN2VmLTQ0NjU1MWE2YjVmNyIsImYyOGExZjUwLWY2ZTctNDU3MS04MThiLTZhMTJmMmFmNmI2YyIsIjdiZTQ0YzhhLWFkYWYtNGUyYS04NGQ2LWFiMjY0OWUwOGExMyIsIjc0OTVmZGM0LTM0YzQtNGQxNS1hMjg5LTk4Nzg4Y2UzOTlmZCIsImIwZjU0NjYxLTJkNzQtNGM1MC1hZmEzLTFlYzgwM2YxMmVmZSIsImIxYmUxYzNlLWI2NWQtNGYxOS04NDI3LWY2ZmEwZDk3ZmViOSIsImFhZjQzMjM2LTBjMGQtNGQ1Zi04ODNhLTY5NTUzODJhYzA4MSIsIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsIjc0ZWY5NzViLTY2MDUtNDBhZi1hNWQyLWI5NTM5ZDgzNjM1MyIsIjRkNmFjMTRmLTM0NTMtNDFkMC1iZWY5LWEzZTBjNTY5NzczYSIsImU2ZDFhMjNhLWRhMTEtNGJlNC05NTcwLWJlZmM4NmQwNjdhNyIsImE5ZWE4OTk2LTEyMmYtNGM3NC05NTIwLThlZGNkMTkyODI2YyIsIjhhYzNmYzY0LTZlY2EtNDJlYS05ZTY5LTU5ZjRjN2I2MGViMiIsIjI5MjMyY2RmLTkzMjMtNDJmZC1hZGUyLTFkMDk3YWYzZTRkZSIsImZkZDdhNzUxLWI2MGItNDQ0YS05ODRjLTAyNjUyZmU4ZmExYyIsIjExNjQ4NTk3LTkyNmMtNGNmMy05YzM2LWJjZWJiMGJhOGRjYyIsIjk2NjcwN2QwLTMyNjktNDcyNy05YmUyLThjM2ExMGYxOWI5ZCIsIjVjNGY5ZGNkLTQ3ZGMtNGNmNy04YzlhLTllNDIwN2NiZmM5MSIsImJhZjM3YjNhLTYxMGUtNDVkYS05ZTYyLWQ5ZDFlNWU4OTE0YiIsImYyZWY5OTJjLTNhZmItNDZiOS1iN2NmLWExMjZlZTc0YzQ1MSIsIjY5MDkxMjQ2LTIwZTgtNGE1Ni1hYTRkLTA2NjA3NWIyYTdhOCIsIjNhMmM2MmRiLTUzMTgtNDIwZC04ZDc0LTIzYWZmZWU1ZDlkNSIsIjBmOTcxZWVhLTQxZWItNDU2OS1hNzFlLTU3YmI4YTNlZmYxZSIsIjliODk1ZDkyLTJjZDMtNDRjNy05ZDAyLWE2YWMyZDVlYTVjMyIsIjY0NGVmNDc4LWUyOGYtNGUyOC1iOWRjLTNmZGRlOWFhMGIxZiIsImU4NjExYWI4LWMxODktNDZlOC05NGUxLTYwMjEzYWIxZjgxNCIsIjE5NGFlNGNiLWIxMjYtNDBiMi1iZDViLTYwOTFiMzgwOTc3ZCIsImUzOTczYmRmLTQ5ODctNDlhZS04MzdhLWJhOGUyMzFjNzI4NiIsImZlOTMwYmU3LTVlNjItNDdkYi05MWFmLTk4YzNhNDlhMzhiMSIsIjNkNzYyYzVhLTFiNmMtNDkzZi04NDNlLTU1YTNiNDI5MjNkNCIsIjE3MzE1Nzk3LTEwMmQtNDBiNC05M2UwLTQzMjA2MmNhY2ExOCIsIjA5NjRiYjVlLTliZGItNGQ3Yi1hYzI5LTU4ZTc5NDg2MmE0MCIsImQzN2M4YmVkLTA3MTEtNDQxNy1iYTM4LWI0YWJlNjZjZTRjMiIsImViMWY0YThkLTI0M2EtNDFmMC05ZmJkLWM3Y2RmNmM1ZWY3YyIsImJlMmY0NWExLTQ1N2QtNDJhZi1hMDY3LTZlYzFmYTYzYmM0NSIsIjZlNTkxMDY1LTliYWQtNDNlZC05MGYzLWU5NDI0MzY2ZDJmMCIsIjcyOTgyN2UzLTljMTQtNDlmNy1iYjFiLTk2MDhmMTU2YmJiOCIsImYwMjNmZDgxLWE2MzctNGI1Ni05NWZkLTc5MWFjMDIyNjAzMyIsIjc1OTQxMDA5LTkxNWEtNDg2OS1hYmU3LTY5MWJmZjE4Mjc5ZSIsImM0ZTM5YmQ5LTExMDAtNDZkMy04YzY1LWZiMTYwZGEwMDcxZiIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfY2MiOlsiQ1AxIl0sInhtc19pZHJlbCI6IjEgMiIsInhtc19zc20iOiIxIiwieG1zX3N0Ijp7InN1YiI6IndCWG1wMVpFX1doX3Rya0NZT3Z5TENfWThXMXFsMmVsb1BDUGZIc1N4MFUifSwieG1zX3RjZHQiOjE0NzY2NzYxOTh9.b7HlTx6j0Dx568uqZvzMz4LlAVwC4b1ppfZbZj6WwioK7hS76fm0mtqX9CiHvB_0dAn_s9B7DcUGP7pZ_luZmUqyb4K2EESyRoL1wYFXB3sII3LlQZtLgz-dFdocPFXfVOvP9d24LFV_My0syLsCJpun3pXKWDKFIMoBKm0zKIhXm6Y3wBNcNjEU6bcLFWogDp2SZ43LwABKz6bv6Z5V7kFUjPQgDo_auURuXRYZeR9txF_mhBr4tMH6bncWYEdgBb_aouNs73CdrB9XUmCsUxtiIKGJPi_YC7j89sw-T7q2UDaaUlCuX1qtxydL2aQB8yAl0FKpH4ii7tlnMu1niw")
			    .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
			    .PUT(HttpRequest.BodyPublishers.ofFile(Paths.get("/opt/files/arquivo_atualizado.docx")))
			    //.PUT(HttpRequest.BodyPublishers.ofFile(Paths.get("C:\\Users\\vinic\\OneDrive\\Documentos\\arquivo_atualizado.docx")))
			    .build();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

        System.out.println("Resposta da API TESTE: " + response.body());
		return "Retorno";
    }
	
    public void atualizarArquivo(String conteudo) {
        try {
            // Abrir o arquivo existente
            //FileInputStream fis = new FileInputStream("C:\\Users\\vinic\\OneDrive\\Documentos\\arquivo.docx");
        	FileInputStream fis = new FileInputStream("/opt/files/arquivo.docx");
            XWPFDocument document = new XWPFDocument(fis);

          
            // Criar um novo parágrafo com texto
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun().setText(conteudo);

            // Salvar o arquivo atualizado
            //FileOutputStream fos = new FileOutputStream("C:\\Users\\vinic\\OneDrive\\Documentos\\arquivo_atualizado.docx");
            FileOutputStream fos = new FileOutputStream("/opt/files/arquivo_atualizado.docx");
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
