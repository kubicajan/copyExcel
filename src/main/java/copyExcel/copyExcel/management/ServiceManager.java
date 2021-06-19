package copyExcel.copyExcel.management;

import copyExcel.copyExcel.models.*;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcelImpl;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;
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

    private final WriteToExcelImpl writeToExcel;

    private final String OPEN_FILE = "source.xlsx";


    @EventListener(ApplicationReadyEvent.class)
    public void process() {

        Map<SheetSpecifics, ArrayList<FYResult>> results;

        Set<String> sheetNamesForReading = Set.of(
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


        List<Coordinate> coordinates = new ArrayList<>();
        String readFromFile = "source.xlsx";
        coordinates.add(new Coordinate("AP12", "BA91"));
        coordinates.add(new Coordinate("AP1067", "BA1070"));

        SourceFileSpecification sourceFile = SourceFileSpecification
                .builder()
                .fileName(readFromFile)
                .coordinates(coordinates)
                .sheets(sheetNamesForReading)
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

        processMe(sourceFile, destinationFile);

    }

    private void processMe(SourceFileSpecification sourceFile, DestinationFileSpecification destinationFile) {
        Map<SheetSpecifics, ArrayList<FYResult>> results;

        try {

            results = readSourceFile(sourceFile);
            writeToDestinationFile(results, destinationFile);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void writeToDestinationFile(Map<SheetSpecifics, ArrayList<FYResult>> results, DestinationFileSpecification destinationFile) throws IOException {
        FileInputStream file;
        String filename = destinationFile.getFileName();

        file = new FileInputStream(new File(filename));
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        writeToExcel.init(results, workbook);
        writeToExcel.writeRegularly(destinationFile.getRegularSheets(), destinationFile.getRegularCoordinate());
        writeToExcel.writeTransposed(destinationFile.getTransposedSheets(), destinationFile.getTransposedCoordinate());

        file.close();
        FileOutputStream outputStream = new FileOutputStream(filename);
        workbook.write(outputStream);
        workbook.close();



    }

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
