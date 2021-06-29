package copyExcel.copyExcel.models;

import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Set;

@Data
@NoArgsConstructor
public class Request {

    //sourceFileSpecification
    private String sourceFileName;
    private List<Coordinate> readFromCoordinates;
    private Set<String> readFromSheets;

    //destinationFileSpecification
    private String destinationFileName;
    private Coordinate transposedCoordinateToStartWritingAt;
    private Coordinate regularCoordinateToStartWritingAt;
    private Set<String> transposedSheetNamesToWriteTo;
    private Set<String> regularSheetsNamesToWriteTo;


}
