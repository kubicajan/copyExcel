package copyExcel.copyExcel.services.impl;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.services.ReadFromExcel;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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

    final private String READ_FROM_FILE = "source.xlsx";

    final private String START_FROM_CELL_COORDINATES = "AP12";
    final private String STOP_AT_CELL_COORDINATE = "BA91";

    final private String ADDITIONAL_START_CELL_COORDINATES = "AP1067";
    final private String ADDITIONAL_STOP_CELL_COORDINATES = "BA1070";

    final private String MEASURE = "FY21_Actuals";


    /**
     * To not load hidden sheets, the ones that we are interested in are in the set. They are named the same as OPCO cells,
     * which are used to filter while writing to file.
     */
    private final Set<String> ACCEPTED_SHEET_NAMES = Set.of(
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
    private String standardizedSheetName;

    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private ArrayList<FYResult> results;

    private CellAddress startCellAddress;
    private CellAddress stopCellAddress;

    private CellAddress additionalStartCellAddress;
    private CellAddress additionalStopCellAddress;


    public Map<SheetSpecifics, ArrayList<FYResult>> process() {
        log.info("Starting to read from... " + READ_FROM_FILE);
        init();
        openFile();
        return (allResults);
    }

    private void init() {
        allResults = new HashMap<>();

        startCellAddress = new CellAddress(START_FROM_CELL_COORDINATES);
        stopCellAddress = new CellAddress(STOP_AT_CELL_COORDINATE);

        additionalStartCellAddress = new CellAddress(ADDITIONAL_START_CELL_COORDINATES);
        additionalStopCellAddress = new CellAddress(ADDITIONAL_STOP_CELL_COORDINATES);
    }


    /**
     * Method opens the file and initiates reading by calling readFile();
     */
    private void openFile() {
        try {
            FileInputStream file = new FileInputStream(new File(READ_FROM_FILE));
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            readFile(workbook);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * Method iterates through all sheets. Data from sheets that passed the filter are saved through method saveData().
     *
     * @param workbook current workbook
     */
    private void readFile(XSSFWorkbook workbook) {
        String tmpSheetName;

        for (Sheet tmpSheet : workbook) {
            results = new ArrayList<>();
            tmpSheetName = tmpSheet.getSheetName().toUpperCase().replace("-", "_");
            if (!ACCEPTED_SHEET_NAMES.contains(tmpSheetName)) {
                continue;
            }
            standardizedSheetName = tmpSheetName;
            sheet = tmpSheet;

            saveData(sheet.getRow(startCellAddress.getRow()), startCellAddress, stopCellAddress);
            saveData(sheet.getRow(additionalStartCellAddress.getRow()), additionalStartCellAddress, additionalStopCellAddress);
        }
    }

    /**
     * Helper method, saves all data between stop cells
     *
     * @param tmpRow       starting row
     * @param startAddress starting cell address
     * @param stopAddress  stop cell address
     */
    private void saveData(Row tmpRow, CellAddress startAddress, CellAddress stopAddress) {
        while (tmpRow.getRowNum() <= stopAddress.getRow()) {
            results.add(buildFYResult(tmpRow, startAddress.getColumn()));
            tmpRow = sheet.getRow(tmpRow.getRowNum() + 1);
        }
        SheetSpecifics sheetSpecifics = new SheetSpecifics(MEASURE, standardizedSheetName);
        allResults.put(sheetSpecifics, results);
    }

    private FYResult buildFYResult(Row row, int address) {
        return FYResult.builder()
                .january(evaluateCell(row, address))
                .february(evaluateCell(row, address + 1))
                .march(evaluateCell(row, address + 2))
                .april(evaluateCell(row, address + 3))
                .may(evaluateCell(row, address + 4))
                .june(evaluateCell(row, address + 5))
                .july(evaluateCell(row, address + 6))
                .august(evaluateCell(row, address + 7))
                .september(evaluateCell(row, address + 8))
                .october(evaluateCell(row, address + 9))
                .november(evaluateCell(row, address + 10))
                .december(evaluateCell(row, address + 11))
                .build();
    }

    private String evaluateCell(Row row, int address) {
        Cell cell = row.getCell(address);
        if (cell.getCellType() == CellType.FORMULA) {
            return cell.getCachedFormulaResultType().toString();
        } else {
            return cell.toString();
        }
    }
}
