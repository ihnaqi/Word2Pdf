
// import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
// import java.io.OutputStreamWriter;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
// import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlOptions;

public class App {
    public static void main(String[] args) {
        try {
            // Read Word document
            FileInputStream fs = new FileInputStream("example.docx");
            XWPFDocument document = new XWPFDocument(fs);
            fs.close();

            // Convert Word document to XML
            XmlOptions options = new XmlOptions();
            options.setSaveOuter();
            FileOutputStream xmlFileOut = new FileOutputStream("example.xml");
            // document.write(xmlFileOut, options);
            document.write(xmlFileOut);
            xmlFileOut.close();

            // Convert XML to PDF
            FileInputStream xmlFileIn = new FileInputStream("example.xml");
            XmlObject xml = XmlObject.Factory.parse(xmlFileIn);
            XWPFDocument documentFromXml = new XWPFDocument(xml.newInputStream());
            PdfOptions pdfOptions = PdfOptions.create();
            FileOutputStream pdfFileOut = new FileOutputStream("example.pdf");
            PdfConverter.getInstance().convert(documentFromXml, pdfFileOut, pdfOptions);
            pdfFileOut.close();
            xmlFileIn.close();

        } catch (IOException | XmlException e) {
            e.printStackTrace();
        }
    }

    // public static void main(String[] args) {
    // try {
    // // Load the Word file
    // FileInputStream fis = new FileInputStream("example.docx");
    // XWPFDocument doc = new XWPFDocument(fis);

    // // Create a new XML file
    // FileOutputStream fos = new FileOutputStream("output.xml");
    // BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(fos));

    // // Extract the text from the Word file
    // XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
    // String text = extractor.getText();

    // // Write the text to the XML file
    // writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    // writer.write("<document>\n");
    // writer.write(text);
    // writer.write("</document>");

    // // Close the files
    // writer.close();
    // fos.close();
    // fis.close();

    // System.out.println("Conversion complete.");

    // } catch (Exception ex) {
    // System.out.println("Error: " + ex.getMessage());
    // }
    // }

}
