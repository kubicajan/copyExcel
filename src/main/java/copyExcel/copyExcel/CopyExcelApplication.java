package copyExcel.copyExcel;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import copyExcel.copyExcel.management.ServiceManager;
import copyExcel.copyExcel.models.Request;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class CopyExcelApplication {
    private static String DATA_FILE = "./src/main/resources/data.json";

    public static void main(String[] args) {
        List<Request> requests = loadConfigurations();
        ServiceManager serviceManager = new ServiceManager(new ReadFromExcel(), new WriteToExcel());

        if (!requests.isEmpty()) {
            for (Request request : requests) {
                serviceManager.process(request);
            }
            return;
        }
        log.warn("No configurations were found.");
    }

    private static List<Request> loadConfigurations() {
        ObjectMapper objectMapper = new ObjectMapper();
        List<Request> requests = new ArrayList<>();

        try {
            requests = objectMapper.readValue(
                    new File(DATA_FILE),
                    new TypeReference<>() {
                    });

        } catch (IOException e) {
            e.printStackTrace();
        }
        return requests;
    }
}
