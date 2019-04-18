import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class HTMLGenerator {

    public static void main(String[] args) {
        HTMLGenerator htmlGenerator = new HTMLGenerator();
        htmlGenerator.parse("");
    }
    public void parse(String fileName) {
        BufferedWriter writer;
        try {
            File file = new File("src/main/resources/test.xlsx");
            String name = file.getName();
            FileInputStream excelFile = new FileInputStream(file);
            Iterator<Row> iterator = null;
            Row row = null;
            int firstrownum = 0;
            int lastrownum = 0;
            if(name.toLowerCase().endsWith("xls")) {
                HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(file));
                HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
                row = hssfSheet.getRow(hssfSheet.getFirstRowNum());
                firstrownum = hssfSheet.getFirstRowNum();
                lastrownum = hssfSheet.getLastRowNum();
                iterator = hssfSheet.iterator();
            } else if(name.toLowerCase().endsWith("xlsx")) {
                XSSFWorkbook xssfWorkbook = new XSSFWorkbook(excelFile);
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
                xssfSheet.getHeader();
                row = xssfSheet.getRow(xssfSheet.getFirstRowNum());
                firstrownum = xssfSheet.getFirstRowNum();
                lastrownum = xssfSheet.getLastRowNum();
                iterator = xssfSheet.iterator();
            }

//            File tempFile = File.createTempFile(name + '.', "htm", new File(file.getParent()));
//            FileOutputStream fileOutputStream = new FileOutputStream(new File(file.getParent(), name + ".htm"));
            writer = new BufferedWriter(new FileWriter(new File(file.getParent(), name.substring(0, name.indexOf(".")) + ".htm")));
            writer.append("<!DOCTYPE html><html><head><title>");
            writer.write(name);
            writer.write("</title></head>");
            writer.write("\n");
            writer.write("<body><table>");

            Iterator<Cell> firstrowcells = row.cellIterator();
            writer.write("\n");
            writer.write("<thead><tr>");
            while(firstrowcells.hasNext()) {
                Cell cell = firstrowcells.next();
                writer.write("<th>");
                writer.write(cell.toString());
                writer.write("</th>");
            }
            writer.write("</tr></thead>");
            writer.write("\n");

            writer.write("<tbody>");
            while(iterator.hasNext()) {
                writer.write("\n");
                writer.write("<tr>");
                Row currentrow = iterator.next();
                int rownum = currentrow.getRowNum();
                if (rownum != firstrownum) {
                    Iterator<Cell> iterator1 = currentrow.cellIterator();
                    while(iterator1.hasNext()) {
                        Cell currentCell = iterator1.next();
                        writer.write("<td>");
                        writer.write(currentCell.toString());
                        writer.write("</td>");
                    }
                    writer.write("</tr>");
                }
            }
            writer.write("</table></body></html>");
            writer.close();
        } catch(Exception e) {
            System.out.println(e.getMessage());
        }

    }
    //test
}
