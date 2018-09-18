package sot.task.wordtask;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import sot.bean.ProcessResult;
import sot.bean.exception.ExcelProcessingException;

import java.io.IOException;

/**
 * <内容说明>
 *
 * @author Monan
 *         created on 2018/8/30 10:45
 */
public abstract class Data2WordBaseTask<T> {
    private Data2WordBaseTask() {
    }

    public abstract ProcessResult<T> execute() throws ExcelProcessingException, Exception;

//    word - > map
//    word -> bean

    /**
     * create a new Excel(2007) process Task
     *
     * @param document     XML Document with Task Configuration
     * @param targets      A String array which storing package locations of bean or full reference of a bean
     * @param workBook2007 Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Data2WordBaseTask(Document document, String[] targets, Workbook workBook2007) throws IOException {
        this.workBook2007 = workBook2007;
        this.document = document;
        this.targets = targets;
    }

    /**
     * create a new Excel(2007) process Task
     *
     * @param taskName     specific which task to read from configuration.
     * @param targets      A String array which storing package locations of bean or full reference of a bean
     * @param workBook2007 Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Data2WordBaseTask(String taskName, String[] targets, Workbook workBook2007) throws IOException {
        this.workBook2007 = workBook2007;
        this.taskName = taskName;
        this.targets = targets;
    }

    /**
     * create a new Excel(2007) process Task
     *
     * @param taskName     specific which task to read from configuration.
     * @param targets      A String array which storing package locations of bean or full reference of a bean
     * @param workBook2007 Excel workBook,work as Data Source.
     * @param sheetName    specific which sheet name that to process,process all available sheet if given <code>Null</code>,which request task mapping in
     *                     configuration.
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Data2WordBaseTask(String taskName, String[] targets, Workbook workBook2007, String sheetName) throws IOException {
        this.workBook2007 = workBook2007;
        this.taskName = taskName;
        this.targets = targets;
        this.sheetName = sheetName;
    }


    /**
     * create a new Excel process Task
     *
     * @param document     XML Document with Task Configuration
     * @param targets      A String array which storing package locations of bean or full reference of a bean
     * @param workBook2003 Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Data2WordBaseTask(Document document, String[] targets, HSSFWorkbook workBook2003) throws IOException {
        this.workBook2007 = workBook2003;
        this.document = document;
        this.targets = targets;
    }

    /**
     * create a new Excel process Task
     *
     * @param taskName     specific which task to read from configuration.
     * @param targets      A String array which storing package locations of bean or full reference of a bean
     * @param workBook2003 Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Data2WordBaseTask(String taskName, String[] targets, HSSFWorkbook workBook2003) throws IOException {
        this.workBook2007 = workBook2003;
        this.taskName = taskName;
        this.targets = targets;
    }

    /**
     * create a new Excel process Task
     *
     * @param taskName     specific which task to read from configuration.
     * @param targets      A String array which storing package locations of bean or full reference of a bean
     * @param workBook2003 Excel workBook2003,work as Data Source.
     * @param sheetName    specific which sheet name that to process,process all available sheet if given <code>Null</code>,which request task mapping in
     *                     configuration.
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Data2WordBaseTask(String taskName, String[] targets, HSSFWorkbook workBook2003, String sheetName) throws IOException {
        this.workBook2007 = workBook2003;
        this.taskName = taskName;
        this.targets = targets;
        this.sheetName = sheetName;
    }

    /**
     * Excel sheet Data 2003 version
     */
    protected HSSFWorkbook workBook2003;

    /**
     * Excel sheet Data 2007 version
     */
    protected Workbook workBook2007;
    /**
     * A String array which storing package locations of bean or full reference of a bean
     */
    protected String[] targets;

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
