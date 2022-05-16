package poi.word;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

import java.io.FileInputStream;

public class WordReader2 {
    public static void main(String[] args) throws Exception {


        FileInputStream fileInputStream = new FileInputStream("1.docx");


        XWPFWordExtractor wordExtractor = new XWPFWordExtractor(OPCPackage.open(fileInputStream));
        String para = wordExtractor.getText();

//        CTR ctr = CTR.Factory.parse(obj.toString());
        System.out.println(para);

    }


}
