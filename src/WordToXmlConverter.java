import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.Fop;
import org.apache.fop.apps.MimeConstants;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import org.xml.sax.helpers.DefaultHandler;

public class WordToXmlConverter {

   public static void main(String[] args) throws Exception {

      // Load the Word file
      XWPFDocument doc = new XWPFDocument(new FileInputStream("example.docx"));

      // Convert the document to XML
      Document xmlDocument = convertToXml(doc);

      // Write the XML file to disk
      TransformerFactory transformerFactory = TransformerFactory.newInstance();
      Transformer transformer = transformerFactory.newTransformer();
      DOMSource source = new DOMSource(xmlDocument);
      StreamResult result = new StreamResult(new FileOutputStream("example.xml"));
      transformer.transform(source, result);

      System.out.println("XML Conversion complete.");

      // XML Conversion is complete so now we need to convert the xml into a pdf
      convertPDF();
   }

   private static Document convertToXml(XWPFDocument doc) {

      // Create a new XML document
      Document xmlDocument = XmlUtils.createDocument();

      // Create the root element
      Element root = xmlDocument.createElement("document");
      xmlDocument.appendChild(root);

      // Convert each paragraph to XML
      for (int i = 0; i < doc.getParagraphs().size(); i++) {
         XWPFParagraph paragraph = doc.getParagraphs().get(i);
         Element element;
         if (paragraph.getStyle() != null && paragraph.getStyle().startsWith("Heading")) {
            // This is a heading paragraph, so wrap it in an h1 or h2 element
            int level = Integer.parseInt(paragraph.getStyle().substring("Heading".length()).trim());
            element = XmlUtils.createElement(xmlDocument, "h" + level);
         } else {
            // This is a regular paragraph, so use a p element
            element = XmlUtils.createElement(xmlDocument, "p");
         }
         element.setTextContent(paragraph.getText());
         root.appendChild(element);
      }

      // Convert each table to XML
      for (int i = 0; i < doc.getTables().size(); i++) {
         Element table = XmlUtils.createElement(xmlDocument, "table");
         root.appendChild(table);
         for (int j = 0; j < doc.getTables().get(i).getRows().size(); j++) {
            Element tr = XmlUtils.createElement(xmlDocument, "tr");
            table.appendChild(tr);
            for (int k = 0; k < doc.getTables().get(i).getRows().get(j).getTableCells().size(); k++) {
               Element td = XmlUtils.createElement(xmlDocument, "td");
               td.setTextContent(doc.getTables().get(i).getRows().get(j).getTableCells().get(k).getText());
               tr.appendChild(td);
            }
         }
      }

      return xmlDocument;
   }

   public static void convertPDF() {
      try {
         File xmlFile = new File("example.xml");
         StreamSource xmlSource = new StreamSource(xmlFile);

         // Step 2: Load the XSLT file as a stream source
         File xsltFile = new File("example.xslt");
         StreamSource xsltSource = new StreamSource(xsltFile);

         // Step 3: Create a transformer factory and transform the XML source using the
         // XSLT source
         TransformerFactory transformerFactory = TransformerFactory.newInstance();
         Transformer transformer = transformerFactory.newTransformer(xsltSource);
         SAXResult intermediateResult = new SAXResult(new DefaultHandler());

         // Step 4: Transform the XML file using the XSLT file and the PDF handler
         transformer.transform(xmlSource, intermediateResult);

         // Step 5: Generate the final PDF file using Apache FOP
         FopFactory fopFactory = FopFactory.newInstance(new File(".").toURI());
         FOUserAgent foUserAgent = fopFactory.newFOUserAgent();
         FileOutputStream pdfOutput = new FileOutputStream(new File("example.pdf"));
         Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, foUserAgent, pdfOutput);
         transformer = transformerFactory.newTransformer();
         SAXResult finalResult = new SAXResult(fop.getDefaultHandler());
         transformer.transform(new StreamSource(new File("intermediate.fo")), finalResult);

         System.out.println("PDF generated successfully.");

      } catch (Exception e) {
         e.printStackTrace();
      }
   }
}