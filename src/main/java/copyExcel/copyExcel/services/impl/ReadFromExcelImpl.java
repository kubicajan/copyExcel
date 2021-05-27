package copyExcel.copyExcel.services.impl;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

@Component
@RequiredArgsConstructor
@Slf4j
public class ReadFromExcelImpl implements ReadFromExcel {

    //reading from file
    final private String READ_FROM_FILE = "202104_MMR FY2021_EUR.xlsx";

    final private String START_FROM_CELL_COORDINATES = "AP12";
    final private String STOP_AT_CELL_COORDINATE = "BA91";

    final private String ADDITIONAL_START_CELL_COORDINATES = "AP1067";
    final private String ADDITIONAL_STOP_CELL_COORDINATES = "BA1070";

    private XSSFSheet sheet;

    //results from all sheets
    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private ArrayList<FYResult> results = new ArrayList<>();


    private final WriteToExcel writeToExcel;

    @EventListener(ApplicationReadyEvent.class)
    public void process() {
        log.info("Starting to read from... " + READ_FROM_FILE);
        readFromFirstExcel();
        writeToExcel.process(allResults);
    }

    private void readFromFirstExcel() {
        try {
            //opening the file
            FileInputStream file = new FileInputStream(new File("../" + READ_FROM_FILE));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            sheet = workbook.getSheetAt(2);

            //define border cells
            CellAddress startCellAddress = new CellAddress(START_FROM_CELL_COORDINATES);
            CellAddress stopCellAddress = new CellAddress(STOP_AT_CELL_COORDINATE);
            CellAddress additionalStartCellAddress = new CellAddress(ADDITIONAL_START_CELL_COORDINATES);
            CellAddress additionalStopCellAddress = new CellAddress(ADDITIONAL_STOP_CELL_COORDINATES);

            //initiating collections to save the cell values to
            allResults = new HashMap<>();
            results = new ArrayList<>();

            //save the data
            saveData(sheet.getRow(startCellAddress.getRow()), startCellAddress, stopCellAddress);
            saveData(sheet.getRow(additionalStartCellAddress.getRow()), additionalStartCellAddress, additionalStopCellAddress);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void saveData(Row tmpRow, CellAddress startAddress, CellAddress stopAddress) {

        //while the current row number is lower than the final row, read and save cell values
        while (tmpRow.getRowNum() <= stopAddress.getRow()) {
            results.add(buildFYResult(tmpRow, startAddress.getColumn()));

            //set next row as the current one and continue loop
            tmpRow = sheet.getRow(tmpRow.getRowNum() + 1);
        }

        SheetSpecifics sheetSpecifics = new SheetSpecifics("FY20_Actuals", "YL_CZ");
        allResults.put(sheetSpecifics, results);
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
