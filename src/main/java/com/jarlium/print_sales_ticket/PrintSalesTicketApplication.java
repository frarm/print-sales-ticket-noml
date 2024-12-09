package com.jarlium.print_sales_ticket;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PrintSalesTicketApplication {

	public static void main(String[] args) throws Exception {

		XWPFDocument document = new XWPFDocument(
				new FileInputStream("D:\\Java\\print-sales-ticket-noml\\input-output\\TicketPlantilla.docx"));

		String productName = args[0];
		String clientNames = args[1];
		String district = args[2];

		for (XWPFParagraph paragraph : document.getParagraphs()) {
			XmlCursor cursor = paragraph.getCTP().newCursor();
			cursor.selectPath(
					"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//*/w:txbxContent/w:p/w:r");

			List<XmlObject> ctrsintxtbx = new ArrayList<XmlObject>();

			while (cursor.hasNextSelection()) {
				cursor.toNextSelection();
				XmlObject obj = cursor.getObject();
				ctrsintxtbx.add(obj);
			}
			for (XmlObject obj : ctrsintxtbx) {
				CTR ctr = CTR.Factory.parse(obj.xmlText());
				XWPFRun bufferrun = new XWPFRun(ctr, (IRunBody) paragraph);
				String text = bufferrun.getText(0);
				if (text != null && text.contains("PRUEBA")) {
					text = text.replace("PRUEBA", productName);
					bufferrun.setText(text, 0);
				} else if (text != null && text.contains("Client")) {
					text = text.replace("Client", clientNames);
					bufferrun.setText(text, 0);
				} else if (text != null && text.contains("District")) {
					text = text.replace("District", district);
					bufferrun.setText(text, 0);
				}
				obj.set(bufferrun.getCTR());
			}
		}

		FileOutputStream out = new FileOutputStream("D:\\Java\\print-sales-ticket-noml\\input-output\\Ticket.docx");
		document.write(out);
		out.close();
		document.close();
	}

}
