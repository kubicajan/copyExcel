package copyExcel.copyExcel.services.impl;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.services.MergeExcels;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Component
@RequiredArgsConstructor
@Slf4j
public class MergeExcelsImpl implements MergeExcels {

    //reading from file
    final private String READ_FROM_FILE = "202104_MMR FY2021_EUR.xlsx";

    final private String START_READING_FROM_CELL_COORDINATES = "AP12";
    final private String STOP_READING_AT_CELL_COORDINATE = "BA91";

    //writing to files
    final private String WRITE_TO_FILE = "MMR FY2020_EUR.xlsx";
    final private String START_WRITING_FROM_CELL_COORDINATES = "D2";

    final private int MEASURE_CELL_POSITION_IN_ROW = 0;
    final private int OPCO_CELL_POSITION_IN_ROW = 1;


    private CellAddress startCellAddress = new CellAddress(START_WRITING_FROM_CELL_COORDINATES);

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    //results from one sheet
    private List<FYResult> results;
    //results from all sheets
    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private int rowCounter;
    private int firstFound;
    private SheetSpecifics sheetSpecifics;


    @EventListener(ApplicationReadyEvent.class)
    public void process() {
        merge();
    }

    private void merge() {
        log.info("Starting to read from... " + READ_FROM_FILE);
        readFromFirstExcel();

        log.info("Reading successful!");
        log.info("Writing to... " + WRITE_TO_FILE);
        openSecondExcel();

        log.info("Writing successful!");
        log.info("Process finished.");
    }

    // loop through all results saved in list when reading the first file (only one sheet per call)
    // rows and columns needed to be switched, so :
    // counter adds to the column number every time there is a new result
    // filterRow() sends information about new line
    private void writeToSecondExcel(ArrayList<FYResult> resultList) {
        int counter = 0;
        firstFound = -1;
        for (FYResult result : resultList) {
            createCell(result.getJanuary(), counter, filterRow());
            createCell(result.getFebruary(), counter, filterRow());
            createCell(result.getMarch(), counter, filterRow());
            createCell(result.getApril(), counter, filterRow());
            createCell(result.getMay(), counter, filterRow());
            createCell(result.getJune(), counter, filterRow());
            createCell(result.getJuly(), counter, filterRow());
            createCell(result.getAugust(), counter, filterRow());
            createCell(result.getSeptember(), counter, filterRow());
            createCell(result.getOctober(), counter, filterRow());
            createCell(result.getNovember(), counter, filterRow());
            createCell(result.getDecember(), counter, filterRow());
            counter++;
            rowCounter = firstFound;
        }
    }

    private void openSecondExcel() {
        try {
            //open the file for writing
            FileInputStream file = new FileInputStream(new File("../" + WRITE_TO_FILE));
            workbook = new XSSFWorkbook(file);
            sheet = workbook.getSheetAt(23);

            for (Map.Entry<SheetSpecifics, ArrayList<FYResult>> resultList : allResults.entrySet()
            ) {
                sheetSpecifics = resultList.getKey();
                writeToSecondExcel(resultList.getValue());
            }
            file.close();

            //write to the file
            FileOutputStream outputStream = new FileOutputStream("../" + WRITE_TO_FILE);
            workbook.write(outputStream);
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //creating cells and setting their values
    private void createCell(String result, int counter, Row tmpRow) {
        int column = startCellAddress.getColumn();
        Cell cell = tmpRow.createCell(column + counter);
        cell.setCellValue(result);
    }

    // function returns a new row that matches the filters.
    // When new list is being processed, firstFound starts with value -1 and the function iterates until it founds a
    // row that suits the filters.
    // Then it sets this row as base point (firstFound = rowCounter) and until the end of the list it uses this value
    // to calculate the appropriate rows.
    private Row filterRow() {
        Row tmpRow = null;
        String measureCell = "";
        String OPCOCell = "";

        // when first row was found it no longer needs to look for others as the excel file is sorted
        if (firstFound != -1) {
            tmpRow = sheet.getRow(startCellAddress.getRow() + firstFound + (rowCounter - firstFound));
            rowCounter++;
            return tmpRow;
        }

        //filtering until the start of a batch of rows is found
        while (!measureCell.equals(sheetSpecifics.getMeasure()) ||
                !OPCOCell.equals(sheetSpecifics.getOpco())
        ) {
            firstFound = rowCounter;
            tmpRow = sheet.getRow(startCellAddress.getRow() + rowCounter);
            measureCell = tmpRow.getCell(MEASURE_CELL_POSITION_IN_ROW).toString();
            OPCOCell = tmpRow.getCell(OPCO_CELL_POSITION_IN_ROW).toString();
            rowCounter++;
        }
        return tmpRow;
    }


    private void readFromFirstExcel() {
        try {
            //opening the file
            FileInputStream file = new FileInputStream(new File("../" + READ_FROM_FILE));
            workbook = new XSSFWorkbook(file);
            sheet = workbook.getSheetAt(2);

            //setting boundaries to know the scope of reading
            CellAddress startCellAddress = new CellAddress(START_READING_FROM_CELL_COORDINATES);
            CellAddress stopCellAddress = new CellAddress(STOP_READING_AT_CELL_COORDINATE);
            Row tmpRow = sheet.getRow(startCellAddress.getRow());

            //initiating array list to save the cell values to
            ArrayList<FYResult> results = new ArrayList<>();
            allResults = new HashMap<>();

            //while the current row number is lower than the final row, read and save cell values
            while (tmpRow.getRowNum() <= stopCellAddress.getRow()) {
                results.add(buildFYResult(tmpRow, startCellAddress.getColumn()));

                //set next row as the current one and continue loop
                tmpRow = sheet.getRow(tmpRow.getRowNum() + 1);
            }
            SheetSpecifics sheetSpecifics = new SheetSpecifics("FY20_Actuals", "YL_CZ");
            SheetSpecifics sheetSpecifics2 = new SheetSpecifics("FY20_Actuals", "YL_ES");
            allResults.put(sheetSpecifics, results);
            allResults.put(sheetSpecifics2, results);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private FYResult buildFYResult(Row row, int address) {
        return FYResult.builder()
                .january(row.getCell(address).toString())
                .february(row.getCell(address + 1).toString())
                .march(row.getCell(address + 2).toString())
                .april(row.getCell(address + 3).toString())
                .may(row.getCell(address + 4).toString())
                .june(row.getCell(address + 5).toString())
                .july(row.getCell(address + 6).toString())
                .august(row.getCell(address + 7).toString())
                .september(row.getCell(address + 8).toString())
                .october(row.getCell(address + 9).toString())
                .november(row.getCell(address + 10).toString())
                .december(row.getCell(address + 11).toString())
                .build();
    }
}
