package copyExcel.copyExcel.models;

import lombok.Builder;
import lombok.Data;

import java.util.Set;

@Data
@Builder
public class DestinationFileSpecification {
    private String fileName;
    private Coordinate transposedCoordinate;
    private Coordinate regularCoordinate;
    private Set<String> transposedSheets;
    private Set<String> regularSheets;
}
