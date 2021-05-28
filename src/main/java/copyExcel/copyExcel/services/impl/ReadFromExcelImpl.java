package copyExcel.copyExcel.services.impl;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
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
import java.util.Set;

@Component
@RequiredArgsConstructor
@Slf4j
public class ReadFromExcelImpl implements ReadFromExcel {

    final private String READ_FROM_FILE = "202104_MMR FY2021_EUR.xlsx";

    final private String START_FROM_CELL_COORDINATES = "AP12";
    final private String STOP_AT_CELL_COORDINATE = "BA91";

    final private String ADDITIONAL_START_CELL_COORDINATES = "AP1067";
    final private String ADDITIONAL_STOP_CELL_COORDINATES = "BA1070";

    final private String MEASURE = "FY21_Actuals";

    //there is a lot of hidden sheets, the ones that we are interested in are named the same as  OPCO cells
    //from the second file
    private static final Set<String> ACCEPTED_SHEET_NAMES = Set.of(
            "YL_BX",
            "YL_CZ",
            "YL_DE",
            "YL_ES",
            "YL_FR",
            "YL_HU",
            "YL_IT",
            "YL_PL",
            "YL_RO",
            "YL_RU",
            "YL_TR",
            "YL_UK");

    private Sheet sheet;
    private String sheetName;

    //results from all sheets
    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    //result from single sheet
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
            //define border cells
            CellAddress startCellAddress = new CellAddress(START_FROM_CELL_COORDINATES);
            CellAddress stopCellAddress = new CellAddress(STOP_AT_CELL_COORDINATE);
            CellAddress additionalStartCellAddress = new CellAddress(ADDITIONAL_START_CELL_COORDINATES);
            CellAddress additionalStopCellAddress = new CellAddress(ADDITIONAL_STOP_CELL_COORDINATES);

            //opening the file
            FileInputStream file = new FileInputStream(new File("../" + READ_FROM_FILE));
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            allResults = new HashMap<>();

            //iterate through all sheets
            for (Sheet tmpSheet : workbook) {
                String tmpSheetName = tmpSheet.getSheetName().toUpperCase().replace("-", "_");
                //check if sheet name is in the set of accepted names. If not, skip.
                if (!ACCEPTED_SHEET_NAMES.contains(tmpSheetName)) {
                    continue;
                }
                sheetName = tmpSheetName;
                sheet = tmpSheet;

                //initiating collections to save the cell values to
                results = new ArrayList<>();

                //save the data
                saveData(sheet.getRow(startCellAddress.getRow()), startCellAddress, stopCellAddress);
                saveData(sheet.getRow(additionalStartCellAddress.getRow()), additionalStartCellAddress, additionalStopCellAddress);
            }
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
        SheetSpecifics sheetSpecifics = new SheetSpecifics(MEASURE, sheetName);
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
