package sot.bean;

import sot.bean.sheetdata.impl.SheetData;
import sot.wordutil.logger.ExcelUtilLogger;

import java.util.Map;

/**
 * Process Result
 *
 * @author Monan
 *         created on 2018/8/30 14:35
 */
public final class ProcessResult<T> {
    private ProcessResult() {
    }

    public ProcessResult(ExcelUtilLogger excelUtilLogger, Map<String, SheetData<T>> sheetDatas) {
        this.excelUtilLogger = excelUtilLogger;
        this.sheetDatas = sheetDatas;
    }

    /**
     * 错误日志记录器
     */
    private ExcelUtilLogger excelUtilLogger;
    /**
     * SheetData
     */
    private Map<String, SheetData<T>> sheetDatas;

    /**
     * 获取任务日志记录器,主要是非致命信息,如:在Excel中找不到某个值
     *
     * @return 任务日志记录器
     */
    public ExcelUtilLogger getExcelUtilLogger() {
        return excelUtilLogger;
    }

    /**
     * Gets SheetData.
     *
     * @return Value of SheetData.
     */
    public Map<String, SheetData<T>> getSheetData() {
        return sheetDatas;
    }
}
