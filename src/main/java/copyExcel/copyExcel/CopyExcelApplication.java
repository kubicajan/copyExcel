package copyExcel.copyExcel;

import copyExcel.copyExcel.management.ServiceManager;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;

public class CopyExcelApplication {
    public static void main(String[] args) {
        ServiceManager serviceManager = new ServiceManager(new ReadFromExcel(), new WriteToExcel());
        serviceManager.process();
    }
}
