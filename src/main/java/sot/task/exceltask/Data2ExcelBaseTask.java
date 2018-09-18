package sot.task.exceltask;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import sot.bean.exception.ExcelProcessingException;
import sot.bean.sheetdata.impl.SheetData;

import java.io.IOException;

/**
 * <内容说明>
 *
 * @author Monan
 *         created on 2018/8/30 10:45
 */
public abstract class Data2ExcelBaseTask {
    private Data2ExcelBaseTask() {
    }

    public abstract Workbook execute() throws ExcelProcessingException, Exception;

    /**
     * Create a Data to Excel(2007) Task
     *
     * @param data         data
     * @param template2007 template
     * @throws IOException
     */
    public Data2ExcelBaseTask(SheetData data, Workbook template2007) throws IOException {
        this.template2007 = template2007;
        this.data = data;
    }

    /**
     * Create a Data to Excel(2003) Task
     *
     * @param data         data
     * @param template2003 template
     * @throws IOException
     */
    public Data2ExcelBaseTask(SheetData data, HSSFWorkbook template2003) throws IOException {
        this.template2003 = template2003;
        this.data = data;
    }

    /**
     * Create a Data to Excel(2007) Task
     *
     * @param data         data
     * @param template2007 template
     * @throws IOException
     */
    public Data2ExcelBaseTask(SheetData data, Workbook template2007, String sheetName) throws IOException {
        this.sheetName = sheetName;
        this.template2007 = template2007;
        this.data = data;
    }

    /**
     * Create a Data to Excel(2003) Task
     *
     * @param data         data
     * @param template2003 template
     * @throws IOException
     */
    public Data2ExcelBaseTask(SheetData data, HSSFWorkbook template2003, String sheetName) throws IOException {
        this.sheetName = sheetName;
        this.template2003 = template2003;
        this.data = data;
    }


    /**
     * Create a Data to Excel(2007) Task
     *
     * @param data     data
     * @param document Task Configuration
     * @param taskName Task Name
     * @throws IOException
     */
    public Data2ExcelBaseTask(SheetData data, Document document, String taskName) throws IOException {
        this.data = data;
        this.document = document;
        this.taskName = taskName;
    }

    /**
     * Template,Excel 2003 version
     */
    protected HSSFWorkbook template2003;

    /**
     * Template,Excel 2007 version
     */
    protected Workbook template2007;

    /**
     * data
     */
    protected SheetData data;

    /**
     * Task Name
     */
    protected String taskName;

    /**
     * targets sheet that need to process
     */
    protected String sheetName;

    /**
     * XML Document with Task Configuration
     */
    protected Document document;
}
