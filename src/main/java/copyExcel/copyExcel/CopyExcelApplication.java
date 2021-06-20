package copyExcel.copyExcel;

import copyExcel.copyExcel.management.ServiceManager;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class CopyExcelApplication {
    public static void main(String[] args) {
        ServiceManager serviceManager = new ServiceManager(new ReadFromExcel(), new WriteToExcel());

        //to catch runtime exceptions and log them
        try {
            serviceManager.process();
        } catch (Exception e) {
            log.error("Something went wrong.", e);
        }
    }
}
