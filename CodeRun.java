package xlreader;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;


public class CodeRun {

/* String userWorkingDirectory = System.getProperty("user.dir");
        String pathSeparator = System.getProperty("file.separator");
        String filepath = userWorkingDirectory +  pathSeparator + "Desktop" + pathSeparator + "ATMECS" + pathSeparator + "Book1.txt";*/
    public static void main(String[] args) {
       
       readExcelFile();

    }


    private static void readExcelFile() {
        List lstStudents;

        try {
            FileInputStream excelFile = new FileInputStream(new File("C:\\Users\\DELL\\Desktop\\ATMECS\\Book1.xlsx"));
            XSSFWorkbook xssfworkbook = new XSSFWorkbook(excelFile);

            Sheet sheet = xssfworkbook.getSheetAt(0);
            Iterator rows = sheet.iterator();

            lstStudents = new ArrayList();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = (Row) rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator cellsInRow = currentRow.iterator();

                ReadExcel stu = new ReadExcel();

                int cellIndex = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = (Cell) cellsInRow.next();

                    if (cellIndex == 0) {
                        stu.setName(currentCell.getStringCellValue());
                    } else if (cellIndex == 1) { // Name
                        stu.setAge((int) currentCell.getNumericCellValue());
                    } else if (cellIndex == 2) {
                        stu.setMarks((int) currentCell.getNumericCellValue());

                    }

                    cellIndex++;
                }

                lstStudents.add(stu);
            }

            // Close WorkBook
            xssfworkbook.close();


        } catch (IOException e) {
            throw new RuntimeException("FAIL! -> message = " + e.getMessage());
        }


        JSONObject object = new JSONObject();
        object.put("Personal details", lstStudents);
        BufferedWriter bufferedWriter = null;
        try {
            bufferedWriter = new BufferedWriter(new FileWriter("Excel_Data.json"));
            bufferedWriter.write(object.toJSONString());
            bufferedWriter.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }


}



