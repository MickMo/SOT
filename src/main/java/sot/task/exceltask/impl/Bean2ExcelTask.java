package sot.task.exceltask.impl;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import sot.bean.exception.ExcelProcessingException;
import sot.bean.sheetdata.impl.SheetData;
import sot.task.exceltask.Data2ExcelBaseTask;

import java.io.IOException;

/**
 * @author Monan
 *         created on 2018/7/27 11:31
 */
public class Bean2ExcelTask extends Data2ExcelBaseTask {

    /**
     * 执行任务
     *
     * @return 处理完成的Excel 工作簿
     * @throws ExcelProcessingException 当处理过程中遇到异常时,如超过错误限制,将会立即抛出ExcelProcessingException;
     * @throws Exception                读取XML Mapping中遇到异常
     */
    @Override
    public Workbook execute() throws ExcelProcessingException, Exception {
        try {
            return null;
        } finally {

        }
    }

    /**
     * 映射信息储存在模板中(多sheet)
     *
     * @param data
     * @param template2007
     * @throws IOException
     */
    public Bean2ExcelTask(SheetData data, Workbook template2007) throws IOException {
        super(data, template2007);
    }

    /**
     * 映射信息储存在模板中(多sheet)
     *
     * @param data
     * @param template2003
     * @throws IOException
     */
    public Bean2ExcelTask(SheetData data, HSSFWorkbook template2003) throws IOException {
        super(data, template2003);
    }

    /**
     * 映射信息储存在模板中
     *
     * @param data
     * @param template2007
     * @param sheetName
     * @throws IOException
     */
    public Bean2ExcelTask(SheetData data, Workbook template2007, String sheetName) throws IOException {
        super(data, template2007, sheetName);
    }

    /**
     * 映射信息储存在模板中
     *
     * @param data
     * @param template2003
     * @param sheetName
     * @throws IOException
     */
    public Bean2ExcelTask(SheetData data, HSSFWorkbook template2003, String sheetName) throws IOException {
        super(data, template2003, sheetName);
    }

    /**
     * 映射信息储存在配置文件中
     *
     * @param data
     * @param document
     * @param taskName
     * @throws IOException
     */
    public Bean2ExcelTask(SheetData data, Document document, String taskName) throws IOException {
        super(data, document, taskName);
    }


}
