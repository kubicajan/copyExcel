package copyExcel.copyExcel;

import copyExcel.copyExcel.management.ServiceManager;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcelImpl;

public class CopyExcelApplication {
    public static void main(String[] args) {
        ServiceManager serviceManager = new ServiceManager(new ReadFromExcel(), new WriteToExcelImpl());
        serviceManager.process();
    }
}
