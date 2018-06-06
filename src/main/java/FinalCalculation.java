import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Arrays;
import java.util.List;

public class FinalCalculation {
    List<String> categories = Arrays.asList("Bank", "NBFI", "Pharmaceuticals");

    void executeFinalCalculation(XSSFWorkbook workbook) {
        XSSFSheet finalSheet = workbook.getSheet("Index&Vol");
        XSSFSheet mCap = workbook.getSheet("MCap");
        XSSFSheet adjDiv = workbook.getSheet("AdjDiv");
        XSSFSheet adjRight = workbook.getSheet("AdjRight");

        for (int i = 1; finalSheet.getRow(i * 8 - 1) != null
                && finalSheet.getRow(i * 8 - 1).getCell(0) != null
                && !finalSheet.getRow(i * 8 - 1).getCell(0).getStringCellValue().equalsIgnoreCase(""); i++) {
            String category = finalSheet.getRow(i * 8 - 1).getCell(0).getStringCellValue();
            finalSheet.getRow((i - 1) * 8 + 1).getCell(0).setCellValue(category + " Market Cap");

            XSSFRow marketCapRow = finalSheet.getRow((i - 1) * 8 + 1);
            for (int capRowCell = 2; capRowCell < marketCapRow.getPhysicalNumberOfCells(); capRowCell++) {
                Double cellValue = 0.0;
                for (int rowNo = 1; rowNo < 356; rowNo++) {
                    if (mCap.getRow(rowNo).getCell(1).getStringCellValue().equalsIgnoreCase(category)) {
                        cellValue += mCap.getRow(rowNo).getCell(capRowCell).getNumericCellValue();
                    }
                }
                marketCapRow.getCell(capRowCell).setCellValue(cellValue);
            }


            XSSFRow entryRow = finalSheet.getRow((i - 1) * 8 + 3);
            for (int entryRowCell = 2; entryRowCell < entryRow.getPhysicalNumberOfCells(); entryRowCell++) {
                Integer entryCnt = 0;
                for (int rowNo = 1; rowNo < 356; rowNo++) {
                    if (mCap.getRow(rowNo).getCell(1).getStringCellValue().equalsIgnoreCase(category)) {
                        if (mCap.getRow(rowNo).getCell(entryRowCell).getNumericCellValue() == 0) {
                            if (mCap.getRow(rowNo).getCell(entryRowCell + 1) != null &&
                                    mCap.getRow(rowNo).getCell(entryRowCell + 1).getNumericCellValue() != 0) {
                                entryCnt++;
                            }
                        }
                    }
                }
                entryRow.getCell(entryRowCell).setCellValue(entryCnt);
            }


            XSSFRow cashDividendRow = finalSheet.getRow((i - 1) * 8 + 4);
            for (int dividentRowCell = 2; dividentRowCell < cashDividendRow.getPhysicalNumberOfCells(); dividentRowCell++) {
                Double cellValue = 0.0;
                for (int rowNo = 1; rowNo < 356; rowNo++) {
                    if (adjDiv.getRow(rowNo).getCell(1).getStringCellValue().equalsIgnoreCase(category)) {
                        cellValue += adjDiv.getRow(rowNo).getCell(dividentRowCell).getNumericCellValue();
                    }
                }
                cashDividendRow.getCell(dividentRowCell).setCellValue(cellValue);
            }


            XSSFRow rightShareRow = finalSheet.getRow((i - 1) * 8 + 5);
            for (int rightShareCell = 2; rightShareCell < rightShareRow.getPhysicalNumberOfCells(); rightShareCell++) {
                Double cellValue = 0.0;
                for (int rowNo = 1; rowNo < 356; rowNo++) {
                    if (adjRight.getRow(rowNo).getCell(1).getStringCellValue().equals(category)) {
                        cellValue += adjRight.getRow(rowNo).getCell(rightShareCell).getNumericCellValue();
                    }
                }
                rightShareRow.getCell(rightShareCell).setCellValue(cellValue);
            }


            XSSFRow adjustmentRow = finalSheet.getRow((i - 1) * 8 + 6);
            for (int adjRowCell = 2; adjRowCell < adjustmentRow.getPhysicalNumberOfCells(); adjRowCell++) {
                Double cellValue = finalSheet.getRow((i - 1) * 8 + 5).getCell(adjRowCell).getNumericCellValue() +
                        finalSheet.getRow((i - 1) * 8 + 3).getCell(adjRowCell).getNumericCellValue() -
                        finalSheet.getRow((i - 1) * 8 + 4).getCell(adjRowCell).getNumericCellValue();
                adjustmentRow.getCell(adjRowCell).setCellValue(cellValue);
            }


            XSSFRow divisorRow = finalSheet.getRow((i - 1) * 8 + 2);

            for (int divisorRowCell = 3; divisorRowCell < divisorRow.getPhysicalNumberOfCells(); divisorRowCell++) {
                Double AW19 = finalSheet.getRow((i - 1) * 8 + 2).getCell(divisorRowCell - 1).getNumericCellValue();
                Double AW18 = finalSheet.getRow((i - 1) * 8 + 1).getCell(divisorRowCell - 1).getNumericCellValue();
                Double AW23 = finalSheet.getRow((i - 1) * 8 + 6).getCell(divisorRowCell - 1).getNumericCellValue();
                Double cellValue = AW19 * ((AW18 + AW23) / AW18);

                divisorRow.getCell(divisorRowCell).setCellValue(cellValue);
            }

            XSSFRow quotientRow = finalSheet.getRow(i * 8 - 1);
            for (int quotientCell = 3; quotientCell < quotientRow.getPhysicalNumberOfCells(); quotientCell++) {
                Double quotientValue = finalSheet.getRow((i - 1) * 8 + 1).getCell(quotientCell).getNumericCellValue()
                        / finalSheet.getRow((i - 1) * 8 + 2).getCell(quotientCell).getNumericCellValue();
                quotientRow.getCell(quotientCell).setCellValue(quotientValue);
            }
        }
        System.out.println("Final sheet done.");
    }


}
