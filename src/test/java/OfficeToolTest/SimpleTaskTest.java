package OfficeToolTest;

import org.apache.poi.ss.usermodel.Workbook;
import sot.bean.ProcessResult;
import sot.bean.exception.ExcelProcessingException;
import sot.bean.sheetdata.impl.SheetData;
import sot.task.exceltask.impl.Excel2MapTask;
import sot.util.fileutil.ExcelFileReadUtil;

import java.io.File;
import java.util.List;
import java.util.Map;

public class SimpleTaskTest {
    public static void main(String[] args) throws Exception, ExcelProcessingException {
        String taskName = "testTask";
        String[] targets = {"OfficeToolTest"};
        File excelFile = new File("src/main/resources/config/testTask.xlsx");
        Workbook workbook = ExcelFileReadUtil.readExcelFile2007(excelFile);
        Excel2MapTask simpleTestTask = new Excel2MapTask(taskName, targets, workbook);
        ProcessResult<Map<String, Object>> result = simpleTestTask.execute();
        Map<String, SheetData<Map<String, Object>>> sheetData = result.getSheetData();
        for (String s : sheetData.keySet()) {
            SheetData<Map<String, Object>> mapSheetData = sheetData.get(s);
            List<Map<String, Object>> rowData = mapSheetData.getRowData();
            for (Map<String, Object> rowDatum : rowData) {
                for (String s1 : rowDatum.keySet()) {
                    System.out.println("Key:" + s1);
                    System.out.println("Value:" + rowDatum.get(s1) );
                }
            }
        }
    }

}

