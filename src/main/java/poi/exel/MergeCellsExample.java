package poi.exel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class MergeCellsExample {
    public static void main(String... args) {

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("merge-cells-example");
        sheet.addMergedRegion(new CellRangeAddress(
                0, //индекс первой строки в отсчете от нуля
                0, //индекс последней строки в отсчете от нуля
                0, //индекс первого столбца в отсчете от нуля
                13  //индекс последнего столбца в отсчете от нуля
        ));

        Row row = sheet.createRow(0);

        Cell cell = row.createCell(0);
        cell.setCellValue("SimpleSolution.dev");



        try (OutputStream fileOut = new FileOutputStream("merge-cells.xls")) {
            workbook.write(fileOut);
            workbook.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
