import DataStructure.TicketSectorDateValue;
import javafx.util.Pair;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class ExcelReader {
    void readExcel() {
        Date nowDate = new Date();
        List<Pair<String, String>> tickerSectorPairList = new ArrayList<Pair<String, String>>();
        String fileName = "/home/insan/Dropbox/Projects/LRGB/src/main/java/Compiled Index.xlsm";
        File excel = new File(fileName);

        TicketSectorDateValue priceSheetValue = null;
        try {
            FileInputStream fs = new FileInputStream(excel);
            XSSFWorkbook myworkbook = null;
            try {
                myworkbook = new XSSFWorkbook(fs);
            } catch (IOException e) {
                e.printStackTrace();
            }
            priceSheetValue = getPriceSheetValues(myworkbook);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    TicketSectorDateValue getPriceSheetValues(XSSFWorkbook myworkbook) {
        XSSFSheet sheet = myworkbook.getSheet("Price");
        Iterator<Row> rowIterator = sheet.iterator();
        Row firstRow = rowIterator.next();
        TicketSectorDateValue ticketSectorDateValue = new TicketSectorDateValue();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            try {
                Pair<String, String> tickerSectorPair = new Pair<String, String>(row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue());
                int cellCount = row.getPhysicalNumberOfCells();
                for (int i = 2; i < cellCount; i++) {
                    Date date = firstRow.getCell(i).getDateCellValue();

                    Pair<Pair<String, String>, Date> ticketSectorDatePair = new Pair<Pair<String, String>, Date>(tickerSectorPair, date);

                    Double cellDoubleValue = row.getCell(i).getNumericCellValue();

                    ticketSectorDateValue.linkedHashMap.put(ticketSectorDatePair, cellDoubleValue);
                }
            } catch (Exception e) {
                System.out.println(e.getMessage());
            }
        }
        return ticketSectorDateValue;
    }


    TicketSectorDateValue getShareNoSheetValue(Date today, XSSFWorkbook myworkbook) {
        TicketSectorDateValue ticketSectorDateValue = new TicketSectorDateValue();




        return ticketSectorDateValue;
    }
}
