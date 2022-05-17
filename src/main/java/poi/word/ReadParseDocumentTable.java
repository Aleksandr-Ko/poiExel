package poi.word;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.List;

public class ReadParseDocumentTable {

    public static void main(String[] args) throws IOException {

        String fileName = "src/main/resources/templateWord/1-1.docx";

        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(fileName)))) {

            // это хорошо только для чтения (нужно дорабатывать) для изменения

            /*XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(doc);
            String docText = xwpfWordExtractor.getText();
            System.out.println(docText);*/

            XWPFWordExtractor wordExtractor = new XWPFWordExtractor(doc);
            String docText = wordExtractor.getText();

            System.out.println(docText);





            Iterator<IBodyElement> docElementsIterator = doc.getBodyElementsIterator();

            //Перебрать список и проверить тип элемента таблицы
            while (docElementsIterator.hasNext()) {
                IBodyElement docElement = docElementsIterator.next();
                if ("TABLE".equalsIgnoreCase(docElement.getElementType().name())) {

                    //Получить список таблиц и повторить его
                    List<XWPFTable> xwpfTableList = docElement.getBody().getTables();
                    for (XWPFTable xwpfTable : xwpfTableList) {
                        System.out.println("Total Rows : " + xwpfTable.getNumberOfRows());
                        for (int i = 0; i < xwpfTable.getRows().size(); i++) {
                            for (int j = 0; j < xwpfTable.getRow(i).getTableCells().size(); j++) {
                                System.out.println(xwpfTable.getRow(i).getCell(j).getText());
                            }
                        }
                    }
                }
            }

        }

    }

}
