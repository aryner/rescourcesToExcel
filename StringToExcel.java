import java.util.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.io.BufferedReader;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class StringToExcel{
  public static void main(String[] args) {
    FileInputStream fileIn = null;
    FileOutputStream fileOut = null;
    InputStreamReader reader = null;
    BufferedReader buffReader = null;

    try {
      fileIn = new FileInputStream(args[0]);
      fileOut = new FileOutputStream(args[1]);
      reader = new InputStreamReader(fileIn);
      buffReader = new BufferedReader(reader);

      ArrayList<String> resources = new ArrayList<String>();
      ArrayList<String> resourceNames = new ArrayList<String>();
      int startIndex = -1;
      int endIndex = -1;
      String resource = null;
      String output = buffReader.readLine();
      while(output != null) {
        startIndex = output.indexOf("\">");
        if(startIndex > 0) {
          endIndex = output.indexOf("</");

          resource = output.substring(startIndex+2, endIndex);
          resources.add(resource);

          endIndex = startIndex;
          startIndex = output.indexOf("name=");

          resourceNames.add(output.substring(startIndex+6,endIndex));
        }

        output = buffReader.readLine();
      }

      Workbook workbook = new XSSFWorkbook();
      CreationHelper createHelper = workbook.getCreationHelper();
      Sheet sheet = workbook.createSheet("sheet1");
      Row row = sheet.createRow((short)0);

      Cell cell = row.createCell(0);
      cell.setCellValue(createHelper.createRichTextString("Resource"));
      cell = row.createCell(1);
      cell.setCellValue(createHelper.createRichTextString("English"));
      cell = row.createCell(2);
      cell.setCellValue(createHelper.createRichTextString("French"));

      for(int i=1; i<resources.size(); i++) {
        row = sheet.createRow((short)i);

        cell = row.createCell(0);
        cell.setCellValue(createHelper.createRichTextString(resourceNames.get(i)));

        cell = row.createCell(1);
        cell.setCellValue(createHelper.createRichTextString(resources.get(i)));
      }

      workbook.write(fileOut);
    }
    catch (Exception e) { e.printStackTrace(); }
    finally {
      try {
        buffReader.close();
        reader.close();
        fileIn.close();
        fileOut.close();
      } catch (Exception e) { e.printStackTrace(); }
    }
  }
}
