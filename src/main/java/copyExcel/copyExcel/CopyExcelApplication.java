package copyExcel.copyExcel;

import copyExcel.copyExcel.management.ServiceManager;
import copyExcel.copyExcel.services.ReadFromExcelImpl;
import copyExcel.copyExcel.services.WriteToExcelImpl;

public class CopyExcelApplication {
    public static void main(String[] args) {
        ServiceManager serviceManager = new ServiceManager(new ReadFromExcelImpl(), new WriteToExcelImpl());
        serviceManager.process();
    }
}
