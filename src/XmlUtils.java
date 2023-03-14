import java.io.File;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class XmlUtils {
   public static Document createDocument() {
      try {
         DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
         DocumentBuilder db = dbf.newDocumentBuilder();
         return db.newDocument();
      } catch (Exception e) {
         throw new RuntimeException(e);
      }
   }

   public static Element createElement(Document doc, String name) {
      return doc.createElement(name);
   }

   public static Document parse(File file) throws Exception {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      DocumentBuilder builder = factory.newDocumentBuilder();
      return builder.parse(file);
   }
}
