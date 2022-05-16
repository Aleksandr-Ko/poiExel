package poi.word;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlException;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;


public class WordReader1 {
    public static void main(String[] args) throws IOException, OpenXML4JException, XmlException {

        FileInputStream fileInputStream = new FileInputStream("8.docx");
        // открываем файл и считываем его содержимое в объект XWPFDocument
        XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxFile);
/*
            // считываем верхний колонтитул (херед документа)

//            XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
//            System.out.println(docHeader.getText());
*/
        // печатаем содержимое всех параграфов документа в консоль
        List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
        for (XWPFParagraph p : paragraphs) {
            System.out.println(p.getText());
        }
/*
        // считываем нижний колонтитул (футер документа)
//            XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
//            System.out.println(docFooter.getText());
*/
/*
 System.out.println("_____________________________________");
 // печатаем все содержимое Word файла
 XWPFWordExtractor extractor = new XWPFWordExtractor(docxFile);
 System.out.println(extractor.getText());
 */

    }

}


