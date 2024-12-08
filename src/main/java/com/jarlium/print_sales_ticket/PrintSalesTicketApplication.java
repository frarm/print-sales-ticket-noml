package com.jarlium.print_sales_ticket;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PrintSalesTicketApplication implements CommandLineRunner {

	public static void main(String[] args) {
		SpringApplication.run(PrintSalesTicketApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {
		if (args.length < 3) {
			System.out.println("Por favor, proporciona nombre, dirección y distrito como argumentos.");
			return;
		}
		String nombre = args[0];
		String direccion = args[1];
		String distrito = args[2];
		// Modificar plantilla de Word
		String wordFilePath = "src/main/resources/TicketPlantilla.docx";
		String modifiedWordFilePath = "C:/Users/Frarm/Downloads/Ticket.pdf";
		modificarPlantillaWord(wordFilePath, modifiedWordFilePath, nombre, direccion, distrito);
		// Convertir a PDF
		String pdfFilePath = "modificado.pdf";
		convertirWordAPdf(modifiedWordFilePath, pdfFilePath);
		// Imprimir PDF
		imprimirPdf(pdfFilePath);
	}

	private void modificarPlantillaWord(String inputFilePath, String outputFilePath, String nombreProducto,
			String nombresCliente,
			String distrito) throws IOException {
		try (FileInputStream fis = new FileInputStream(inputFilePath); XWPFDocument document = new XWPFDocument(fis)) {
			for (XWPFParagraph paragraph : document.getParagraphs()) {
				for (XWPFRun run : paragraph.getRuns()) {
					String text = run.getText(0);
					if (text != null) {
						text = text.replace("{nombre}", nombreProducto).replace("{direccion}", nombresCliente).replace(
								"{distrito}",
								distrito);
						run.setText(text, 0);
					}
				}
			}
			try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
				document.write(fos);
			}
		}
	}

	private void convertirWordAPdf(String wordFilePath, String pdfFilePath) throws IOException {
		try (PDDocument pdfDocument = new PDDocument()) {
			PDPage page = new PDPage(PDRectangle.A4);
			pdfDocument.addPage(page);
			try (PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page)) {
				contentStream.beginText();
				contentStream.setFont(PDType1Font.HELVETICA, 12);
				contentStream.newLineAtOffset(100, 700);
				contentStream.showText("Contenido del documento Word convertido a PDF.");
				contentStream.endText();
			}
			pdfDocument.save(pdfFilePath);
		}
	}

	private void imprimirPdf(String pdfFilePath) {
		// Lógica para imprimir el archivo PDF
		System.out.println("Imprimiendo archivo PDF: " + pdfFilePath);
	}
}
