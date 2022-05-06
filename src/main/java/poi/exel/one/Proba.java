package poi.exel.one;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class Proba {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Header header = sheet.getHeader();
//        header.setCenter("Center Header");
//        header.setLeft("Left Header");
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
                HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");


        try (OutputStream fileOut = new FileOutputStream("workbook.xls")) {
            wb.write(fileOut);
        }

    }
}
