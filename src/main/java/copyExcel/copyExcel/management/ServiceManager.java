package copyExcel.copyExcel.management;

import copyExcel.copyExcel.models.*;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Component
@RequiredArgsConstructor
@Slf4j
public class ServiceManager {

    private final ReadFromExcel readFromExcel;

    private final WriteToExcel writeToExcel;

    private final Set<String> SHEET_NAMES_FOR_READING = Set.of(
            "YL-BX",
            "YL-CZ",
            "YL-DE",
            "YL-ES",
            "YL-FR",
            "YL-HU",
            "YL-IT",
            "YL-PL",
            "YL-RO",
            "YL-RU",
            "YL-TR",
            "YL-UK");


    public void process() {
        firstFilePair();
        secondFilePair();
    }


    /**
     * This method takes care of configuring the files, that will be processed.
     * <p>
     * Here the destinations for reading and writing are set, together with the coordinates and
     * deciding which sheets will be used for transposed or regular writing.
     * <p>
     * At the end of the method, the whole process is initiated.
     */
    private void firstFilePair() {
        List<Coordinate> coordinates = new ArrayList<>();
        String readFromFile = "source.xlsx";
        coordinates.add(new Coordinate("AP12", "BA91"));
        coordinates.add(new Coordinate("AP1067", "BA1070"));

        SourceFileSpecification sourceFile = SourceFileSpecification
                .builder()
                .fileName(readFromFile)
                .coordinates(coordinates)
                .sheets(SHEET_NAMES_FOR_READING)
                .build();

        String writeToFile = "MMR FY2020_EUR.xlsx";
        Set<String> transposedSheets = Set.of("MMR_2");
        Coordinate transposedCoordinate = new Coordinate("D2", null);

        Set<String> regularSheets = Set.of("FY21");
        Coordinate regularCoordinate = new Coordinate("F2", null);


        DestinationFileSpecification destinationFile = DestinationFileSpecification
                .builder()
                .fileName(writeToFile)
                .transposedCoordinate(transposedCoordinate)
                .regularCoordinate(regularCoordinate)
                .transposedSheets(transposedSheets)
                .regularSheets(regularSheets)
                .build();

        process(sourceFile, destinationFile);
    }


    /**
     * Very similarly to the firstFilePair() method, this one is also a configuration and initialization
     * of the whole process. If there is a need to work with additional config, another method should be
     * created
     */
    private void secondFilePair() {
        List<Coordinate> coordinates = new ArrayList<>();
        String readFromFile = "source.xlsx";
        coordinates.add(new Coordinate("AP597", "BA676"));
        coordinates.add(new Coordinate("AP482", "BA485"));

        SourceFileSpecification sourceFile = SourceFileSpecification
                .builder()
                .fileName(readFromFile)
                .coordinates(coordinates)
                .sheets(SHEET_NAMES_FOR_READING)
                .build();

        String writeToFile = "MMR_EUR - JPY.xlsx";
        Set<String> transposedSheets = Set.of("MMR_2");
        Coordinate transposedCoordinate = new Coordinate("D2", null);

        Set<String> regularSheets = Set.of("FY21");
        Coordinate regularCoordinate = new Coordinate("F2", null);


        DestinationFileSpecification destinationFile = DestinationFileSpecification
                .builder()
                .fileName(writeToFile)
                .transposedCoordinate(transposedCoordinate)
                .regularCoordinate(regularCoordinate)
                .transposedSheets(transposedSheets)
                .regularSheets(regularSheets)
                .build();

        process(sourceFile, destinationFile);
    }

    /**
     * Method handles calls for reading and then writing data to its respected files
     */
    private void process(SourceFileSpecification sourceFile, DestinationFileSpecification destinationFile) {
        Map<SheetSpecifics, ArrayList<FYResult>> results;

        try {
            log.info("Starting to read from... " + sourceFile.getFileName());
            results = readSourceFile(sourceFile);
            log.info("Reading finished");

            log.info("Starting to write to... " + destinationFile.getFileName());
            writeToDestinationFile(results, destinationFile);
            log.info("Writing finished");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Method opens the source file and initiates reading, then closes it.
     */
    private void writeToDestinationFile(Map<SheetSpecifics, ArrayList<FYResult>> results, DestinationFileSpecification destinationFile) throws IOException {
        FileInputStream file;
        String filename = destinationFile.getFileName();

        file = new FileInputStream(new File(filename));
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        writeToExcel.init(results, workbook);
        writeToExcel.initiateWriting(destinationFile.getRegularSheets(), destinationFile.getRegularCoordinate(), true);
        writeToExcel.initiateWriting(destinationFile.getTransposedSheets(), destinationFile.getTransposedCoordinate(), false);

        file.close();
        FileOutputStream outputStream = new FileOutputStream(filename);
        workbook.write(outputStream);
        workbook.close();
    }


    /**
     * Method opens the source file and initiates reading, then closes it.
     */
    private Map<SheetSpecifics, ArrayList<FYResult>> readSourceFile(SourceFileSpecification sourceFile) throws IOException {
        FileInputStream file;
        String filename = sourceFile.getFileName();
        Map<SheetSpecifics, ArrayList<FYResult>> results;
        XSSFWorkbook workbook = new XSSFWorkbook(filename);

        file = new FileInputStream(new File(filename));
        results = readFromExcel.process(sourceFile, workbook);

        workbook.close();
        file.close();
        return results;
    }
}
