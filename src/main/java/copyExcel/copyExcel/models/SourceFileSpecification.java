package copyExcel.copyExcel.models;

import lombok.Builder;
import lombok.Data;

import java.util.List;
import java.util.Set;

@Data
@Builder
public class SourceFileSpecification {
    private String fileName;
    private List<Coordinate> coordinates;
    private Set<String> sheets;
}
