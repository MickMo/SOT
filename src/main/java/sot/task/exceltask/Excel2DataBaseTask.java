package sot.task.exceltask;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import org.dom4j.Node;
import sot.bean.BeanInfo;
import sot.bean.ProcessResult;
import sot.bean.exception.ExcelProcessingException;
import sot.bean.sheetdata.impl.SheetData;
import sot.wordutil.logger.ExcelUtilLogger;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * <内容说明>
 *
 * @author Monan
 *         created on 2018/8/30 10:45
 */
public abstract class Excel2DataBaseTask<T> {
    private Excel2DataBaseTask() {
    }

    public abstract ProcessResult<T> execute() throws ExcelProcessingException, Exception;

    /**
     * create a new Excel(2007) process Task
     *
     * @param document     XML Document with Task Configuration
     * @param targets      A String array which storing package locations of bean or full reference of a bean
     * @param workBook2007 Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Excel2DataBaseTask(Document document, String[] targets, Workbook workBook2007) throws IOException {
        invalidInitialCheck();
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
    public Excel2DataBaseTask(String taskName, String[] targets, Workbook workBook2007) throws IOException {
        invalidInitialCheck();
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
    public Excel2DataBaseTask(String taskName, String[] targets, Workbook workBook2007, String sheetName) throws IOException {
        invalidInitialCheck();
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
    public Excel2DataBaseTask(Document document, String[] targets, HSSFWorkbook workBook2003) throws IOException {
        invalidInitialCheck();
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
    public Excel2DataBaseTask(String taskName, String[] targets, HSSFWorkbook workBook2003) throws IOException {
        invalidInitialCheck();
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
    public Excel2DataBaseTask(String taskName, String[] targets, HSSFWorkbook workBook2003, String sheetName) throws IOException {
        invalidInitialCheck();
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

    /**
     * Map that store the processed result
     */
    protected Map<String, SheetData<T>> sheetDatas;

    /**
     * Map that store the Mapping information between Bean Object and Excel Column
     */
    protected Map<String, List<BeanInfo>> beanInfosFromXMLMap;

    /**
     * Head Row Number For Each Excel Sheet K-V:SheetName-HeadNum
     */
    protected Map<String, Integer> taskHeaderRowNums;

    /**
     * Process Logger
     */
    protected ExcelUtilLogger excelUtilLogger;

    /**
     * Task Node in XML
     */
    protected Node taskNode = null;


    private void invalidInitialCheck() {
        if (workBook2003 != null && workBook2007 != null) {
            throw new RuntimeException("The Task Could Not be Initialize Twice!");
        }
    }
}
