import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

public class App {
    public static void main(String[] args) {

        try {
            FileInputStream fs = new FileInputStream("example.docx");
            XWPFDocument document = new XWPFDocument(fs);

            Document pdfDoc = new Document();
            PdfWriter.getInstance(pdfDoc, new FileOutputStream("example.pdf"));
            pdfDoc.open();

            for (XWPFParagraph para : document.getParagraphs()) {

                if (para.isPageBreak()) {
                    pdfDoc.newPage();
                }

                int spaceBef = para.getSpacingBeforeLines();
                int spaceAfter = para.getSpacingAfterLines();

                Paragraph newPara = new Paragraph(para.getText());
                System.out.println(newPara);

                newPara.setSpacingBefore(spaceBef);
                newPara.setSpacingAfter(spaceAfter);
                pdfDoc.add(newPara);

            }

            pdfDoc.close();
            fs.close();

        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }

    }
}
