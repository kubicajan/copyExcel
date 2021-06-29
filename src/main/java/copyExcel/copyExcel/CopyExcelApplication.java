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
    public static void main(String[] args) {
        List<Request> requests = loadConfigurations();
        //todo: sometimes a bug happens: https://stackoverflow.com/questions/64116589/error-occurs-while-saving-the-package-the-part-xl-sharedstrings-xml-fail-to-b

        if (!requests.isEmpty()) {
            for (Request request : requests) {
                startThread(request);
            }
            return;
        }
        log.warn("No configurations were found.");
    }

    private static void startThread(Request request) {

    }

    private static List<Request> loadConfigurations() {
        ObjectMapper objectMapper = new ObjectMapper();
        List<Request> requests = new ArrayList<>();

        try {
            requests = objectMapper.readValue(
                    new File("data.json"),
                    new TypeReference<>() {
                    });

        } catch (IOException e) {
            e.printStackTrace();
        }
        return requests;
    }
}
