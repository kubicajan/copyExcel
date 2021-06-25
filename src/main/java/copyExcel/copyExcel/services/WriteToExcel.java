package copyExcel.copyExcel.services;

import copyExcel.copyExcel.models.Coordinate;
import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Map;
import java.util.Set;

@RequiredArgsConstructor
@Slf4j
public class WriteToExcel {

    final private int MEASURE_CELL_POSITION_IN_ROW = 0;
    final private int OPCO_CELL_POSITION_IN_ROW = 1;

    private XSSFWorkbook workbook;
    private Sheet sheet;
    private int rowCounter;

    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private SheetSpecifics sheetSpecifics;
    private CellAddress startCellAddress;


    /**
     * Method initiates the writing to the respected file. It first needs to get the correct sheet, which
     * due to some weird naming in the destination files sometimes contains random spaces. Thus we need
     * to go through every sheet, remove spaces and check, if the name is the name is the one we want.
     *
     * @param regular decides, whether the data should be written transposed or regularly
     */
    public void initiateWriting(Set<String> sheets, Coordinate coordinate, boolean regular) {
        startCellAddress = new CellAddress(coordinate.getBeginCoordinate());
        String tmpSheetName;

        for (Sheet tmpSheet : workbook) {
            tmpSheetName = tmpSheet.getSheetName().replace(" ", "");

            if (sheets.contains(tmpSheetName)) {
                sheet = workbook.getSheet(tmpSheet.getSheetName());
                log.info("Writing to sheet... " + tmpSheetName);


                for (Map.Entry<SheetSpecifics, ArrayList<FYResult>> resultList : allResults.entrySet()) {
                    sheetSpecifics = resultList.getKey();
                    if (regular) {
                        writeIntoSheetRegularly(resultList.getValue());
                    } else {
                        writeIntoSheetTransposed(resultList.getValue());

                    }
                }
            }
        }
    }

    public void init(Map<SheetSpecifics, ArrayList<FYResult>> results, XSSFWorkbook sentWorkbook) {
        allResults = results;
        workbook = sentWorkbook;
    }


    /**
     * Method writes the results into a file, each FYResult represents one row and each parameter
     * of the said FYResult represents one column.
     */
    private void writeIntoSheetRegularly(ArrayList<FYResult> resultList) {
        Row tmpRow;
        int columnNumber;
        int counter = 0;

        for (FYResult result : resultList) {
            tmpRow = sheet.getRow(filterRows().getRowNum() + counter);
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

    /**
     * Method writes the results into a file, each FYResult represents one column and each parameter
     * of the said FYResult represents one row.
     */
    private void writeIntoSheetTransposed(ArrayList<FYResult> resultList) {
        int counter = 0;
        int firstBatchRowNumber;
        Row firstBatchRow;
        int columnNumber;

        for (FYResult result : resultList) {
            rowCounter = 0;
            firstBatchRow = filterRows();
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


    /**
     * Method goes through all rows and looks for the row with Measure and OSCO cells that fit
     * the ones we saved into the SheetSpecifics in allResults map.
     */
    private Row filterRows() {
        Row tmpRow = null;
        String measureCell = "";
        String OPCOCell = "";
        int rowCounter = 0;

        while (!measureCell.equals(sheetSpecifics.getMeasure()) ||
                !OPCOCell.equals(sheetSpecifics.getOpco())) {
            tmpRow = sheet.getRow(startCellAddress.getRow() + rowCounter);
            measureCell = tmpRow.getCell(MEASURE_CELL_POSITION_IN_ROW).toString();
            OPCOCell = tmpRow.getCell(OPCO_CELL_POSITION_IN_ROW).toString();
            rowCounter = rowCounter + 1;
        }
        if (tmpRow != null) {
            return tmpRow;
        }
        throw new IllegalArgumentException("No row fits the criteria");
    }


    /**
     * Helper class for creating classes.
     *
     * @param result       will be written as a value for newly created row
     * @param columnNumber Column number where new cell will be created
     * @param tmpRow       Row where new cell will be created
     */
    private void createCell(Double result, int columnNumber, Row tmpRow) {
        Cell cell = tmpRow.createCell(columnNumber);
        if (result != null) {
            cell.setCellValue(result);
        }
    }

    /**
     * Helper method for horizontal writing, fetches row one higher every time it gets called.
     *
     * @param rowNumber row number
     * @return returns new row
     */
    private Row getNextRowIncrementing(int rowNumber) {
        Row tmpRow = sheet.getRow(rowNumber + rowCounter);
        rowCounter++;
        return tmpRow;
    }

}
