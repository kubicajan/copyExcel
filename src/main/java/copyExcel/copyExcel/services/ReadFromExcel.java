package copyExcel.copyExcel.services;

import copyExcel.copyExcel.models.Coordinate;
import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.models.SourceFileSpecification;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

@RequiredArgsConstructor
@Slf4j
public class ReadFromExcel {

    // measure will probaly need to get changed every year.
    final private String MEASURE = "FY21_Actuals";

    private Set<String> acceptedSheetNames;
    private List<Coordinate> coordinates;
    private Map<SheetSpecifics, ArrayList<FYResult>> allResults;
    private ArrayList<FYResult> results;


    /**
     * Function that gets called from the outside and starts the whole flow
     *
     * @return function returns all lines read from the excel
     */
    public Map<SheetSpecifics, ArrayList<FYResult>> process(SourceFileSpecification sourceFile, XSSFWorkbook workbook) {
        init(sourceFile);
        readFile(workbook);
        return allResults;
    }

    private void init(SourceFileSpecification sourceFile) {
        allResults = new HashMap<>();
        acceptedSheetNames = sourceFile.getSheets();
        coordinates = sourceFile.getCoordinates();
    }


    /**
     * Method goes through sheets and initiates reading according to the coordinates
     */
    private void readFile(XSSFWorkbook workbook) {
        CellAddress startCellAddress;
        CellAddress stopCellAddress;
        Sheet sheet;

        for (String tmpSheetName : acceptedSheetNames) {
            results = new ArrayList<>();
            sheet = workbook.getSheet(tmpSheetName);

            if (sheet == null) {
                log.warn("Sheet " + tmpSheetName + " does not exist, skipping.");
                continue;
            }
            log.info("Reading from sheet " + tmpSheetName);

            for (Coordinate coordinate : coordinates) {
                startCellAddress = new CellAddress(coordinate.getBeginCoordinate());
                stopCellAddress = new CellAddress(coordinate.getEndCoordinate());
                saveData(sheet, startCellAddress, stopCellAddress);
            }
        }
    }

    /**
     * Method start saving lines to a list, which gets saved as a value to map allResults. Its key
     * is sheetSpecifics, which contains measure and opsco values used for identifying the correct
     * lines to write to.
     * <p>
     * It is needed to replace the "_" for "-" in sheet name, which gets saved to opsco property of SheetSpecifics
     * object. We do this because the destination file has opsco fields named the same as the sheets of
     * the source file, just with "-" instead of "_"
     */
    private void saveData(Sheet sheet, CellAddress startAddress, CellAddress stopAddress) {
        Row tmpRow = sheet.getRow(startAddress.getRow());
        String standardizedSheetName = sheet.getSheetName().toUpperCase().replace("-", "_");

        while (tmpRow.getRowNum() <= stopAddress.getRow()) {
            results.add(buildFYResult(tmpRow, startAddress.getColumn()));
            tmpRow = sheet.getRow(tmpRow.getRowNum() + 1);
        }
        SheetSpecifics sheetSpecifics = new SheetSpecifics(MEASURE, standardizedSheetName);
        allResults.put(sheetSpecifics, results);
    }


    /**
     * Helper method for building FYResultq
     */
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

    private double evaluateCell(Row row, int address) {
        Cell cell = row.getCell(address);

        switch (cell.getCellType()) {
            case FORMULA:
                return cell.getNumericCellValue();
            case NUMERIC:
                return Double.parseDouble(cell.toString());
        }
        throw new IllegalArgumentException("Cell " + cell + " containing a formula does not fit any of the possible cases");
    }
}
