import DataStructure.TicketSectorDateValue;
import javafx.util.Pair;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class ExcelReader2 {
    void readExcel() {
        Date nowDate = new Date();
        List<Pair<String, String>> tickerSectorPairList = new ArrayList<Pair<String, String>>();
        String inputFileName = "/home/insan/Dropbox/Projects/LRGB/src/main/java/Compiled Index.xlsm";
        String outputFileName = "/home/insan/Dropbox/Projects/LRGB/src/main/java/output.xlsm";
        File inputExcel = new File(inputFileName);

        FileInputStream fi = null;
        XSSFWorkbook workbookInput = null;
        XSSFWorkbook workbookInput2 = null;
        try {
            Long startTime = System.currentTimeMillis();
            fi = new FileInputStream(inputExcel);
            workbookInput = new XSSFWorkbook(fi);
            Long endTime = System.currentTimeMillis();
            System.out.println("Input file read. Time taken: " +  (endTime - startTime) + " miliseconds.");
            exportSheetsToCSV(workbookInput);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }

    private void exportSheetsToCSV(XSSFWorkbook workbookInput) throws IOException {

        DataFormatter formatter = new DataFormatter();

        for (Sheet sheet : workbookInput) {
            File csvFile = new File("/home/insan/Dropbox/Projects/LRGB/src/main/java/" + sheet.getSheetName() + ".csv");
            csvFile.createNewFile();
            PrintStream out = null;
            try {
                out = new PrintStream(new FileOutputStream(csvFile),
                        true, "UTF-8");
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            for (Row row : sheet) {
                boolean firstCell = true;
                for (Cell cell : row) {
                    if (!firstCell) out.print(',');
                    String text = formatter.formatCellValue(cell);
                    out.print(text);
                    firstCell = false;
                }
                out.println();
            }
        }

    }
}
