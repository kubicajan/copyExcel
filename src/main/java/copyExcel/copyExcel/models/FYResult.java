package copyExcel.copyExcel.models;

import lombok.Builder;
import lombok.Data;

@Builder
@Data
public class FYResult {
    private double january;
    private double february;
    private double march;
    private double april;
    private double may;
    private double june;
    private double july;
    private double august;
    private double september;
    private double october;
    private double november;
    private double december;
}
