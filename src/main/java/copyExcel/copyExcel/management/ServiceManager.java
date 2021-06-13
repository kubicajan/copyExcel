package copyExcel.copyExcel.management;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;
import copyExcel.copyExcel.services.ReadFromExcel;
import copyExcel.copyExcel.services.WriteToExcel;
import lombok.RequiredArgsConstructor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;

@Component
@RequiredArgsConstructor
public class ServiceManager {

    private final ReadFromExcel readFromExcel;

    private final WriteToExcel writeToExcel;


    @EventListener(ApplicationReadyEvent.class)
    public void process() {
        Map<SheetSpecifics, ArrayList<FYResult>> results;

        results = readFromExcel.process();
        writeToExcel.process(results);
    }
}
