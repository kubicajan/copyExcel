package copyExcel.copyExcel.models;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
@AllArgsConstructor
public class SheetSpecifics {
    private String measure;
    private String opco;
}
