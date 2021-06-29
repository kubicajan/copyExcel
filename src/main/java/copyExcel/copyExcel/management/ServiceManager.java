package copyExcel.copyExcel.management;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.Request;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;

@Component
@RequiredArgsConstructor
@Slf4j
public class ServiceManager {

    private final ReadFromExcel readFromExcel;

    private final WriteToExcel writeToExcel;

    /**
     * Method handles calls for reading and then writing data to its respected files
     */
    public void process(Request request) {
        Map<SheetSpecifics, ArrayList<FYResult>> results;

        try {
            log.info("Starting to read from... " + request.getSourceFileName());
            results = readSourceFile(request);
            log.info("Reading from... " + request.getSourceFileName() + "  finished" + "\n");

            log.info("Starting to write to... " + request.getDestinationFileName());
            writeToDestinationFile(results, request);
            log.info("Writing to... " + request.getDestinationFileName() + " finished" + "\n");

        } catch (IOException e) {
            log.error("Something went wrong.", e);
        }
    }

    /**
     * Method opens the source file and initiates reading, then closes it.
     */
    private void writeToDestinationFile(Map<SheetSpecifics, ArrayList<FYResult>> results, Request request) throws IOException {
        String filename = request.getDestinationFileName();
        FileOutputStream outputStream;
        FileInputStream file;
        Workbook workbook;

        file = new FileInputStream(new File(filename));
        workbook = new XSSFWorkbook(file);

        writeToExcel.init(results, workbook);
        writeToExcel.initiateWriting(request.getRegularSheetsNamesToWriteTo(), request.getRegularCoordinateToStartWritingAt(), true);
        writeToExcel.initiateWriting(request.getTransposedSheetNamesToWriteTo(), request.getTransposedCoordinateToStartWritingAt(), false);

        file.close();
        outputStream = new FileOutputStream(filename);
        workbook.write(outputStream);
        workbook.close();
    }


    /**
     * Method opens the source file and initiates reading, then closes it.
     */
    private Map<SheetSpecifics, ArrayList<FYResult>> readSourceFile(Request request) throws IOException {
        Map<SheetSpecifics, ArrayList<FYResult>> results;
        XSSFWorkbook workbook;
        FileInputStream file;
        String filename;


        filename = request.getSourceFileName();
        workbook = new XSSFWorkbook(filename);
        file = new FileInputStream(new File(filename));

        readFromExcel.init(request.getReadFromSheets(), request.getReadFromCoordinates());
        results = readFromExcel.readFile(workbook);

        workbook.close();
        file.close();
        return results;
    }
}
