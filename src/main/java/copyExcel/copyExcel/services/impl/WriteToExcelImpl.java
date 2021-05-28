package copyExcel.copyExcel.services.impl;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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
    final private String START_WRITING_FROM_CELL_COORDINATES = "D2";

    final private int MEASURE_CELL_POSITION_IN_ROW = 0;
    final private int OPCO_CELL_POSITION_IN_ROW = 1;

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private int rowCounter;

    //amount of attributes in FYResult is the same as amount of lines that are skipped when filtering throu rows.
    private int numberOfAttributesInFYResult;

    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private SheetSpecifics sheetSpecifics;
    private CellAddress startCellAddress = new CellAddress(START_WRITING_FROM_CELL_COORDINATES);

    public void process(Map<SheetSpecifics, ArrayList<FYResult>> results) {
        init();
        allResults = results;
        log.info("Reading successful!");
        log.info("Writing to... " + WRITE_TO_FILE);
        openExcel();

        log.info("Writing successful!");
        log.info("Process finished.");
    }

    private void openExcel() {
        try {
            //open the file for writing
            FileInputStream file = new FileInputStream(new File("../" + WRITE_TO_FILE));
            workbook = new XSSFWorkbook(file);
            //todo: change to dynamically found sheetindex
            sheet = workbook.getSheetAt(23);


            for (Map.Entry<SheetSpecifics, ArrayList<FYResult>> resultList : allResults.entrySet()
            ) {
                sheetSpecifics = resultList.getKey();
                writeToExcel(resultList.getValue());
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
            firstBatchRow = filterRows();
            firstBatchRowNumber = firstBatchRow.getRowNum();
            columnNumber = startCellAddress.getColumn() + counter;

            createCell(result.getJanuary(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getFebruary(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getMarch(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getApril(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getMay(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getJune(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getJuly(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getAugust(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getSeptember(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getOctober(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getNovember(), columnNumber, getNextRow(firstBatchRowNumber));
            createCell(result.getDecember(), columnNumber, getNextRow(firstBatchRowNumber));
            counter++;
        }
    }

    private Row filterRows() {
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
            rowCounter = rowCounter + numberOfAttributesInFYResult;
        }
        return tmpRow;
    }

    //creating cells and setting their values
    private void createCell(String result, int columnNumber, Row tmpRow) {
        Cell cell = tmpRow.createCell(columnNumber);
        cell.setCellValue(result);
    }

    private Row getNextRow(int rowNumber) {
        Row tmpRow = sheet.getRow(rowNumber + rowCounter);
        rowCounter++;
        return tmpRow;
    }

}
