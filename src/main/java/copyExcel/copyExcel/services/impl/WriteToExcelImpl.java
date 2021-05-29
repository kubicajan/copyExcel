package copyExcel.copyExcel.services.impl;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;

@NoArgsConstructor
@AllArgsConstructor
@Component
@Slf4j
public class WriteToExcelImpl implements WriteToExcel {

    final private String WRITE_TO_FILE = "MMR FY2020_EUR.xlsx";

    //the result for MMR2 needs to be flipped horizontally, FY21 does not need to be flipped

    final private String MMR_2_SHEET_NAME = "MMR_2";
    final private String FY21_SHEET_NAME = "FY21";

    final private String START_WRITING_FROM_CELL_COORDINATES_MMR_2 = "D2";
    final private String START_WRITING_FROM_CELL_COORDINATES_FY21 = "F2";

    final private int MEASURE_CELL_POSITION_IN_ROW = 0;
    final private int OPCO_CELL_POSITION_IN_ROW = 1;

    private XSSFWorkbook workbook;
    private Sheet sheet;
    private int rowCounter;

    //amount of attributes in FYResult is the same as amount of lines that are skipped when filtering throu rows.
    private int numberOfAttributesInFYResult;
    private int numberOfFyResults;

    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private SheetSpecifics sheetSpecifics;
    private CellAddress startCellAddress;

    public void process(Map<SheetSpecifics, ArrayList<FYResult>> results) {
        allResults = results;
        init();
        log.info("Reading successful!");
        log.info("Writing to... " + WRITE_TO_FILE);
        openExcel();

        log.info("Writing successful!");
        log.info("Process finished.");
    }

    private void writeIntoFY21(ArrayList<FYResult> resultList) {
        Row tmpRow;
        int columnNumber;
        int counter = 0;

        for (FYResult result : resultList) {
            tmpRow = sheet.getRow(filterRows(numberOfFyResults).getRowNum() + counter);
            columnNumber = startCellAddress.getColumn();

            createCell(result.getJanuary(), columnNumber, tmpRow);
            createCell(result.getFebruary(), columnNumber + 1, tmpRow);
            createCell(result.getMarch(), columnNumber + 2, tmpRow);
            createCell(result.getApril(), columnNumber + 3, tmpRow);
            createCell(result.getMay(), columnNumber + 4, tmpRow);
            createCell(result.getJune(), columnNumber + 5, tmpRow);
            createCell(result.getJuly(), columnNumber + 6, tmpRow);
            createCell(result.getAugust(), columnNumber + 7, tmpRow);
            createCell(result.getSeptember(), columnNumber + 8, tmpRow);
            createCell(result.getOctober(), columnNumber + 9, tmpRow);
            createCell(result.getNovember(), columnNumber + 10, tmpRow);
            createCell(result.getDecember(), columnNumber + 11, tmpRow);
            counter++;
        }

    }

    private void decide(ArrayList<FYResult> resultList) {
        switch (sheet.getSheetName()) {
            case MMR_2_SHEET_NAME:
//                startCellAddress = new CellAddress(START_WRITING_FROM_CELL_COORDINATES_MMR_2);
//                writeToExcel(resultList);
                break;
            case FY21_SHEET_NAME:
                startCellAddress = new CellAddress(START_WRITING_FROM_CELL_COORDINATES_FY21);
                writeIntoFY21(resultList);
                break;
        }
    }

    private void openExcel() {
        try {
            //open the file for writing
            FileInputStream file = new FileInputStream(new File("../" + WRITE_TO_FILE));
            workbook = new XSSFWorkbook(file);

            for (Sheet tmpSheet : workbook) {
                if (!tmpSheet.getSheetName().equals(MMR_2_SHEET_NAME) &&
                        !tmpSheet.getSheetName().equals(FY21_SHEET_NAME)
                ) {
                    continue;
                }
                sheet = tmpSheet;
                for (Map.Entry<SheetSpecifics, ArrayList<FYResult>> resultList : allResults.entrySet()
                ) {
                    numberOfFyResults = resultList.getValue().size();
                    sheetSpecifics = resultList.getKey();
                    //decide
                    decide(resultList.getValue());
                }
            }

            file.close();

            //write to the file
            FileOutputStream outputStream = new FileOutputStream("../" + WRITE_TO_FILE);
            workbook.write(outputStream);
            workbook.close();

        } catch (
                IOException e) {
            e.printStackTrace();
        }

    }

    private void init() {
        FYResult r = FYResult.builder().build();
        numberOfAttributesInFYResult = r.getNumberOfAttributes();
    }


    // loop through all results saved in list when reading the first file (only one sheet per call)
    // rows and columns needed to be switched, so :
    // counter adds to the column number every time there is a new result
    // filterRow() sends information about new line
    private void writeToExcel(ArrayList<FYResult> resultList) {
        int counter = 0;
        int firstBatchRowNumber;
        Row firstBatchRow;
        int columnNumber;

        for (FYResult result : resultList) {
            rowCounter = 0;
            firstBatchRow = filterRows(numberOfAttributesInFYResult);
            firstBatchRowNumber = firstBatchRow.getRowNum();
            columnNumber = startCellAddress.getColumn() + counter;

            createCell(result.getJanuary(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getFebruary(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getMarch(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getApril(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getMay(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getJune(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getJuly(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getAugust(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getSeptember(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getOctober(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getNovember(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            createCell(result.getDecember(), columnNumber, getNextRowIncrementing(firstBatchRowNumber));
            counter++;
        }
    }

    private Row filterRows(int skipBy) {
        Row tmpRow = null;
        String measureCell = "";
        String OPCOCell = "";
        int rowCounter = 0;


        while (!measureCell.equals(sheetSpecifics.getMeasure()) ||
                !OPCOCell.equals(sheetSpecifics.getOpco())
        ) {
            tmpRow = sheet.getRow(startCellAddress.getRow() + rowCounter);
            measureCell = tmpRow.getCell(MEASURE_CELL_POSITION_IN_ROW).toString();
            OPCOCell = tmpRow.getCell(OPCO_CELL_POSITION_IN_ROW).toString();
            rowCounter = rowCounter + skipBy;
        }
        return tmpRow;
    }

    //creating cells and setting their values
    private void createCell(String result, int columnNumber, Row tmpRow) {
        Cell cell = tmpRow.createCell(columnNumber);
        cell.setCellValue(result);
    }

    private Row getNextRowIncrementing(int rowNumber) {
        Row tmpRow = sheet.getRow(rowNumber + rowCounter);
        rowCounter++;
        return tmpRow;
    }

}
