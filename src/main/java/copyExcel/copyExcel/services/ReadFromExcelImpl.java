package copyExcel.copyExcel.services;

import copyExcel.copyExcel.models.Coordinate;
import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.models.SourceFileSpecification;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.util.*;

@Component
@RequiredArgsConstructor
@Slf4j
public class ReadFromExcelImpl {

//    final private String READ_FROM_FILE = "source.xlsx";

//    final private String START_FROM_CELL_COORDINATES = "AP12";
//    final private String STOP_AT_CELL_COORDINATE = "BA91";
//
//    final private String ADDITIONAL_START_CELL_COORDINATES = "AP1067";
//    final private String ADDITIONAL_STOP_CELL_COORDINATES = "BA1070";

    // todo: actuals
    final private String MEASURE = "FY21_Actuals";


    /**
     * To not load hidden sheets, the ones that we are interested in are in the set. They are named the same as OPCO cells,
     * which are used to filter while writing to file.
     */
    private Set<String> acceptedSheetNames;

    private Sheet sheet;
    private String standardizedSheetName;
    private List<Coordinate> coordinates;

    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private ArrayList<FYResult> results;

    private CellAddress startCellAddress;
    private CellAddress stopCellAddress;


    public Map<SheetSpecifics, ArrayList<FYResult>> process(SourceFileSpecification sourceFile, XSSFWorkbook workbook) {
        log.info("Starting to read from... " + sourceFile);
        init(sourceFile);
        readFile(workbook);
        return (allResults);
    }

    private void init(SourceFileSpecification sourceFile) {
        allResults = new HashMap<>();
        acceptedSheetNames = sourceFile.getSheets();
        coordinates = sourceFile.getCoordinates();
//        startCellAddress = new CellAddress(sourceFile.getCoordinates().getBeginCoordinate());
//        stopCellAddress = new CellAddress(sourceFile.getCoordinates().getEndCoordinate());
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

            if (acceptedSheetNames.contains(tmpSheetName)) {
                standardizedSheetName = tmpSheetName;
                sheet = tmpSheet;
                for (Coordinate coordinate : coordinates) {
                    startCellAddress = new CellAddress(coordinate.getBeginCoordinate());
                    stopCellAddress = new CellAddress(coordinate.getEndCoordinate());
                    saveData(sheet.getRow(startCellAddress.getRow()), startCellAddress, stopCellAddress);
                }

            }
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
            return evaluateFormulaCell(cell);
        } else {
            return cell.toString();
        }
    }

    private String evaluateFormulaCell(Cell cell) {
        switch (cell.getCachedFormulaResultType()) {
            case _NONE:
                break;
            case NUMERIC:
                return Double.toString(cell.getNumericCellValue());
            case STRING:
                return cell.getStringCellValue();
            case BLANK:
                return "";
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
        }
        throw new IllegalArgumentException("Cell " + cell + " containing a formula has not fit any of the possible cases");
    }
}
