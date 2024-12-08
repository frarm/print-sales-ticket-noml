package com.jarlium.print_sales_ticket;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

public class TextBoxExtractor {

	public static void main(String[] args) throws Exception {
		try (FileInputStream fis = new FileInputStream("src/main/resources/TicketPlantilla.docx")) {
			XWPFDocument document = new XWPFDocument(fis);
			for (XWPFParagraph paragraph : document.getParagraphs()) {
				List<XWPFRun> runs = paragraph.getRuns();
				for (int i = 0; i < runs.size(); i++) {
					XWPFRun run = runs.get(i);
					List<CTDrawing> drawings = getAllDrawings(run);
					for (CTDrawing drawing : drawings) {
						String textBoxContent = getTextBoxContent(drawing);
						System.out.println("Run ID/Position: " + i);
						System.out.println("TextBox Content: " + textBoxContent);
					}
				}
			}
		}
	}

	private static List<CTDrawing> getAllDrawings(XWPFRun run) throws Exception {
		CTR ctR = run.getCTR();
		XmlCursor cursor = ctR.newCursor();
		cursor.selectPath(
				"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:drawing");
		List<CTDrawing> drawings = new ArrayList<CTDrawing>();
		while (cursor.hasNextSelection()) {
			cursor.toNextSelection();
			XmlObject obj = cursor.getObject();
			CTDrawing drawing = CTDrawing.Factory.parse(obj.newInputStream());
			drawings.add(drawing);
		}
		return drawings;
	}

	private static String getTextBoxContent(CTDrawing drawing) {
		StringBuilder result = new StringBuilder();
		XmlCursor cursor = drawing.newCursor();
		cursor.selectPath(
				"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:txbxContent");
		while (cursor.hasNextSelection()) {
			cursor.toNextSelection();
			result.append(cursor.getTextValue());
		}
		return result.toString();
	}
}
