package copyExcel.copyExcel.services;

import copyExcel.copyExcel.models.Coordinate;
import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.Map;
import java.util.Set;

@NoArgsConstructor
@AllArgsConstructor
@Component
@Slf4j
public class WriteToExcelImpl {

    final private String WRITE_TO_FILE = "MMR FY2020_EUR.xlsx";

    // the result for MMR2 needs to be flipped horizontally, FY21 does not need to be flipped
    final private String MMR_2_SHEET_NAME = "MMR_2";
    final private String FY21_SHEET_NAME = "FY21";

    //     private String START_WRITING_FROM_CELL_COORDINATES_MMR_2;
//     = "D2";
//     private String START_WRITING_FROM_CELL_COORDINATES_FY21;
    private String START_WRITING_FROM_CELL_COORDINATES;
//     "F2";

    final private int MEASURE_CELL_POSITION_IN_ROW = 0;
    final private int OPCO_CELL_POSITION_IN_ROW = 1;

    private XSSFWorkbook workbook;
    private Sheet sheet;
    private int rowCounter;

    private Set<String> ACCEPTED_SHEET_NAMES;
//            Set.of(
//            "FY21",
//            "MMR_2"
//    );


    /**
     * the amount of attributes in FYResult is the same as amount of lines that are skipped when filtering through rows (flipped writing)
     */
    private int amountOfAttributesInFYResult;

    /**
     * The amount of results is the same as the amount of lines skipped when filtering rows (normal writing)
     */
    private int amountOfFyResults;

    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private SheetSpecifics sheetSpecifics;
    private CellAddress startCellAddress;

    public void writeRegularly(Set<String> sheets, Coordinate coordinate) {
        ACCEPTED_SHEET_NAMES = sheets;
        startCellAddress = new CellAddress(coordinate.getBeginCoordinate());
        for (String sheetName : sheets
        ) {
            sheet = workbook.getSheet(sheetName);

            for (Map.Entry<SheetSpecifics, ArrayList<FYResult>> resultList : allResults.entrySet()) {
                sheetSpecifics = resultList.getKey();
                writeIntoSheetRegularly(resultList.getValue());
            }
        }

    }

    public void writeTransposed(Set<String> sheets, Coordinate coordinate) {
        ACCEPTED_SHEET_NAMES = sheets;
        startCellAddress = new CellAddress(coordinate.getBeginCoordinate());
        for (String sheetName : sheets
        ) {
            sheet = workbook.getSheet(sheetName);
            for (Map.Entry<SheetSpecifics, ArrayList<FYResult>> resultList : allResults.entrySet()) {
                sheetSpecifics = resultList.getKey();
                writeIntoSheetTransposed(resultList.getValue());
            }
        }

    }

    public void init(Map<SheetSpecifics, ArrayList<FYResult>> results, XSSFWorkbook sentWorkbook) {
        allResults = results;
        workbook = sentWorkbook;

        FYResult r = FYResult.builder().build();
        amountOfAttributesInFYResult = r.getNumberOfAttributes();

        Map.Entry<SheetSpecifics, ArrayList<FYResult>> entry = results.entrySet().iterator().next();
        ArrayList<FYResult> value = entry.getValue();
        amountOfFyResults = value.size();
    }


    /**
     * For each result in resultList, createCell() is called. Parameters sent ensure linear writing, each
     * result represents one row
     */
    private void writeIntoSheetRegularly(ArrayList<FYResult> resultList) {
        Row tmpRow;
        int columnNumber;
        int counter = 0;

        for (FYResult result : resultList) {
            tmpRow = sheet.getRow(filterRows(1).getRowNum() + counter);
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
     * For each result in resultList, createCell() is called. Parameters sent ensure horizontal writing. Each result
     * represents one column.
     */
    //todo: change skipbys
    private void writeIntoSheetTransposed(ArrayList<FYResult> resultList) {
        int counter = 0;
        int firstBatchRowNumber;
        Row firstBatchRow;
        int columnNumber;

        for (FYResult result : resultList) {
            rowCounter = 0;
            firstBatchRow = filterRows(1);
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
     * Method filters through cells in sheet and returns first row that fits the criteria. As the excel file is sorted,
     * this allows us to skip batches of data and return only the first row of the batch.
     *
     * @param skipBy the amount of rows contained in one batch, differs for horizontal and vertical writing.
     * @return first row found that fits criteria.
     */
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


    /**
     * Helper class for creating classes.
     *
     * @param result       will be written as a value for newly created row
     * @param columnNumber Column number where new cell will be created
     * @param tmpRow       Row where new cell will be created
     */
    private void createCell(String result, int columnNumber, Row tmpRow) {
        Cell cell = tmpRow.createCell(columnNumber);
        cell.setCellValue(result);
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
