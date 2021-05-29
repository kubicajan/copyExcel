package copyExcel.copyExcel.services;

import copyExcel.copyExcel.models.FYResult;
import copyExcel.copyExcel.models.SheetSpecifics;

import java.util.ArrayList;
import java.util.Map;

public interface ReadFromExcel {
    Map<SheetSpecifics, ArrayList<FYResult>> process();
}
