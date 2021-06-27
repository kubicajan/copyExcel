package copyExcel.copyExcel;

import copyExcel.copyExcel.management.ServiceManager;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class CopyExcelApplication {
    public static void main(String[] args) {
        ServiceManager serviceManager = new ServiceManager(new ReadFromExcel(), new WriteToExcel());
        ServiceManager serviceManager2 = new ServiceManager(new ReadFromExcel(), new WriteToExcel());

        //todo: sometimes a bug happens: https://stackoverflow.com/questions/64116589/error-occurs-while-saving-the-package-the-part-xl-sharedstrings-xml-fail-to-b

        //todo: check if everything gets loaded correctly.
        
        //to catch runtime exceptions and log them
        try {
            new Thread(serviceManager::firstFilePair).start();
            new Thread(serviceManager2::secondFilePair).start();

        } catch (Exception e) {
            log.error("Something went wrong.", e);
        }
    }
}
