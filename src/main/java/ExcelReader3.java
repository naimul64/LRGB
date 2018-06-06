import DataStructure.TicketSectorDateValue;
import com.microsoft.schemas.office.visio.x2012.main.RowType;
import javafx.util.Pair;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Color;
import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import static java.lang.Double.NaN;

public class ExcelReader3 {
    void readExcel() {
        Date nowDate = new Date();
        List<Pair<String, String>> tickerSectorPairList = new ArrayList<Pair<String, String>>();
        String inputFileName = "/home/insan/Dropbox/Projects/LRGB/src/main/java/output.xlsm";
        File inputExcel = new File(inputFileName);

        FileInputStream fi = null;
        XSSFWorkbook workbookInput = null;
        try {
            Long startTime = System.currentTimeMillis();
            fi = new FileInputStream(inputExcel);
            workbookInput = new XSSFWorkbook(fi);
            Long endTime = System.currentTimeMillis();
            System.out.println("Input file read. Time taken: " + (endTime - startTime) + " miliseconds.");
            createSheetAdjDiv(workbookInput);
            createSheetAdjRight(workbookInput);
            createSheetMCap(workbookInput);
            populateTDivisor(workbookInput);
            populateTReturn(workbookInput);

            fi.close();
            FileOutputStream fo = new FileOutputStream(inputExcel);
            workbookInput.write(fo);
            fo.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }

    void createSheetAdjDiv(XSSFWorkbook workbook) {

        XSSFSheet adjDiv = getSheetWithNameAndBasicCells("AdjDiv", workbook);

        adjDiv.setTabColor(new XSSFColor());
        XSSFSheet shareNo = workbook.getSheet("ShareNumber");
        XSSFSheet cDividend = workbook.getSheet("CDividend");

        int startRow = 1;
        int startColumn = 2;

        for (int rowNo = startRow; rowNo < shareNo.getPhysicalNumberOfRows() &&
                rowNo < cDividend.getPhysicalNumberOfRows(); rowNo++) {
            XSSFRow row = adjDiv.getRow(rowNo);
            for (int cellNo = startColumn; cellNo < shareNo.getRow(rowNo).getPhysicalNumberOfCells() &&
                    cellNo < cDividend.getRow(rowNo).getPhysicalNumberOfCells(); cellNo++) {
                Double cellValue;
                try {
                    cellValue = shareNo.getRow(rowNo).getCell(cellNo).getNumericCellValue() *
                            cDividend.getRow(rowNo).getCell(cellNo).getNumericCellValue();
                } catch (Exception ex) {
                    System.out.println("Row: " + rowNo + " Cell: " + cellNo + '\n' + ex.getMessage());
                    cellValue = 0.0;
                }
                XSSFCell cell = row.createCell(cellNo, CellType.NUMERIC);
                cell.setCellValue(cellValue);
            }
        }
    }

    void createSheetAdjRight(XSSFWorkbook workbook) {
        XSSFSheet adjDiv = getSheetWithNameAndBasicCells("AdjRight", workbook);

        XSSFSheet shareNo = workbook.getSheet("ShareNumber");
        XSSFSheet multiplier = workbook.getSheet("Multiplier");
        XSSFSheet rightP = workbook.getSheet("RightP");

        int startRow = 1;
        int startColumn = 2;

        int rowLimit = Math.max(Math.max(shareNo.getPhysicalNumberOfRows(), multiplier.getPhysicalNumberOfRows()), rightP.getPhysicalNumberOfRows());
        for (int rowNo = startRow; rowNo < rowLimit && rowNo < adjDiv.getPhysicalNumberOfRows(); rowNo++) {
            XSSFRow row = adjDiv.getRow(rowNo);

            for (int cellNo = startColumn; cellNo < adjDiv.getRow(0).getPhysicalNumberOfCells(); cellNo++) {
                Double cellValue;
                try {
                    cellValue = shareNo.getRow(rowNo).getCell(cellNo).getNumericCellValue() *
                            multiplier.getRow(rowNo).getCell(cellNo).getNumericCellValue() *
                            rightP.getRow(rowNo).getCell(cellNo).getNumericCellValue();
                } catch (Exception ex) {
                    System.out.println("Row: " + rowNo + " Cell: " + cellNo + '\n' + getStackTraceAsString(ex));
                    cellValue = 0.0;
                }
                XSSFCell cell = row.createCell(cellNo, CellType.NUMERIC);
                cell.setCellValue(cellValue);
            }
        }
    }

    void createSheetMCap(XSSFWorkbook workbook) {
        XSSFSheet mCap = getSheetWithNameAndBasicCells("MCap", workbook);

        XSSFSheet priceSheet = workbook.getSheet("Price");
        XSSFSheet shareNo = workbook.getSheet("ShareNumber");

        int priceInitial = getInitialForPrice(priceSheet, shareNo);

        int startRow = 1;
        int startColumn = 2;

        int rowLimit = Math.max(shareNo.getPhysicalNumberOfRows(), priceSheet.getPhysicalNumberOfRows());
        for (int rowNo = startRow; rowNo < rowLimit && rowNo < mCap.getPhysicalNumberOfRows(); rowNo++) {
            XSSFRow row = mCap.getRow(rowNo);

            for (int cellNo = startColumn; cellNo < mCap.getRow(0).getPhysicalNumberOfCells(); cellNo++) {
                Double cellValue;
                try {
                    cellValue = priceSheet.getRow(rowNo).getCell(cellNo + priceInitial - startColumn).getNumericCellValue() *
                            shareNo.getRow(rowNo).getCell(cellNo).getNumericCellValue();
                    System.out.println(cellValue);
                } catch (Exception ex) {
                    System.out.println("Row: " + rowNo + " Cell: " + cellNo + '\n' + getStackTraceAsString(ex));
                    cellValue = 0.0;
                }
                XSSFCell cell = row.createCell(cellNo, CellType.NUMERIC);
                cell.setCellValue(cellValue);
            }
        }
    }

    void createTDivisor(XSSFWorkbook workbook) {
        XSSFSheet tDivisor = getSheetWithNameAndBasicCells("TDivisor", workbook);

        XSSFSheet mCap = workbook.getSheet("MCap");
        XSSFSheet adjDiv = workbook.getSheet("AdjDiv");
        XSSFSheet adjRight = workbook.getSheet("AdjRight");

        int startRow = 1;
        int startColumn = 2;


        for (int rowNo = startRow; rowNo < tDivisor.getPhysicalNumberOfRows(); rowNo++) {
            XSSFRow row = tDivisor.getRow(rowNo);

            for (int cellNo = startColumn; cellNo < tDivisor.getRow(0).getPhysicalNumberOfCells(); cellNo++) {
                Double cellValue;
                try {
                    cellValue = 0.0;
                } catch (Exception ex) {
                    System.out.println("Row: " + rowNo + " Cell: " + cellNo + '\n' + getStackTraceAsString(ex));
                    cellValue = 0.0;
                }
                XSSFCell cell = row.createCell(cellNo, CellType.NUMERIC);
                cell.setCellValue(cellValue);
            }
        }
    }

    void populateTDivisor(XSSFWorkbook workbook) {
        XSSFSheet tDivisor = workbook.getSheet("TDivisor");
        XSSFSheet mCap = workbook.getSheet("MCap");
        XSSFSheet adjDiv = workbook.getSheet("AdjDiv");
        XSSFSheet adjRight = workbook.getSheet("AdjRight");

        for (int rowNo = 1; rowNo < tDivisor.getPhysicalNumberOfRows(); rowNo++) {
            XSSFRow tDivisorRow = tDivisor.getRow(rowNo);
            for (int cellNo = 3; cellNo < tDivisorRow.getPhysicalNumberOfCells(); cellNo++) {
                Double cellValue = null;
                try {
                    cellValue = tDivisor.getRow(rowNo).getCell(cellNo - 1).getNumericCellValue() *
                            ((mCap.getRow(rowNo).getCell(cellNo).getNumericCellValue() - adjDiv.getRow(rowNo).getCell(cellNo).getNumericCellValue() + adjRight.getRow(rowNo).getCell(cellNo).getNumericCellValue()) / mCap.getRow(rowNo).getCell(cellNo).getNumericCellValue());
                } catch (Exception ex) {
                    System.out.println("Row: " + rowNo + " Cell: " + cellNo + '\n' + getStackTraceAsString(ex));
                    cellValue = 0.0;
                }

                XSSFCell cell = tDivisorRow.createCell(cellNo, CellType.NUMERIC);
                cell.setCellValue(Double.isNaN(cellValue) ? 0 : cellValue);
            }
        }
    }

    void populateTReturn(XSSFWorkbook workbook) {
        XSSFSheet tReturn = workbook.getSheet("TReturn");
        XSSFSheet tDivisor = workbook.getSheet("TDivisor");
        XSSFSheet mCap = workbook.getSheet("MCap");

        for (int rowNo = 1; rowNo < tReturn.getPhysicalNumberOfRows(); rowNo++) {
            XSSFRow tReturnRow = tReturn.getRow(rowNo);
            for (int cellNo = 3; cellNo < tReturnRow.getPhysicalNumberOfCells(); cellNo++) {
                Double cellValue = null;
                try {
                    cellValue = mCap.getRow(rowNo).getCell(cellNo).getNumericCellValue()/tDivisor.getRow(rowNo).getCell(cellNo - 1).getNumericCellValue();
                } catch (Exception ex) {
                    System.out.println("Row: " + rowNo + " Cell: " + cellNo + '\n' + getStackTraceAsString(ex));
                    cellValue = tReturn.getRow(rowNo).getCell(cellNo - 1).getNumericCellValue();
                }

                XSSFCell cell = tReturnRow.createCell(cellNo, CellType.NUMERIC);
                cell.setCellValue(Double.isNaN(cellValue) || Double.isInfinite(cellValue) ? tReturn.getRow(rowNo).getCell(cellNo - 1).getNumericCellValue() : cellValue);
            }
        }
    }

    int getInitialForPrice(XSSFSheet priceSheet, XSSFSheet shareNo) {
        int priceInitial = -99;
        for (int i = 2; i < priceSheet.getRow(0).getPhysicalNumberOfCells(); i++) {
            if (shareNo.getRow(0).getCell(2).getNumericCellValue()
                    == (priceSheet.getRow(0).getCell(i).getNumericCellValue())) {
                priceInitial = i;
                break;
            }
        }

        return priceInitial;
    }

    XSSFSheet getSheetWithNameAndBasicCells(String sheetName, XSSFWorkbook workbook) {
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }

        sheet.setTabColor(new XSSFColor(Color.GREEN));

        XSSFSheet cDividendSheet = workbook.getSheet("CDividend");

        for (int rowNo = 0; rowNo < cDividendSheet.getPhysicalNumberOfRows(); rowNo++) {
            XSSFRow curSheetRow = sheet.createRow(rowNo);
            XSSFCell cell0 = curSheetRow.createCell(0, CellType.STRING);
            XSSFCell cell1 = curSheetRow.createCell(1, CellType.STRING);

            try {
                cell0.setCellValue(cDividendSheet.getRow(rowNo).getCell(0).getStringCellValue());
                cell1.setCellValue(cDividendSheet.getRow(rowNo).getCell(1).getStringCellValue());
            } catch (Exception ex) {
                System.out.println("Row no: " + rowNo + "\n" + ex.getMessage());
            }
        }


        XSSFRow cDivSheetfirstRow = cDividendSheet.getRow(0);
        XSSFRow currentSheetFirstRow = sheet.getRow(0);
        for (int cellNo = 2; cellNo < cDivSheetfirstRow.getPhysicalNumberOfCells(); cellNo++) {
            XSSFCell curSheetCell = currentSheetFirstRow.createCell(cellNo);
            curSheetCell.setCellValue(cDivSheetfirstRow.getCell(cellNo).getDateCellValue());
        }

        return sheet;
    }


    String getStackTraceAsString(Exception ex) {
        StringWriter sw = new StringWriter();
        PrintWriter pw = new PrintWriter(sw);
        ex.printStackTrace(pw);
        String sStackTrace = sw.toString(); // stack trace as a string
        return sStackTrace;
    }
}
