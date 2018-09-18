package sot.task.exceltask.impl;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import sot.bean.BeanInfo;
import sot.bean.ProcessResult;
import sot.bean.exception.ExcelProcessingException;
import sot.bean.sheetdata.impl.SheetData;
import sot.task.exceltask.Excel2DataBaseTask;
import sot.util.XmlReadUtil;
import sot.util.excelprocessutil.ArrayToBeanArrayUtil;
import sot.util.excelprocessutil.ExcelTo2DArrayUtil;
import sot.util.fileutil.ReadExcelMappingUtil;
import sot.wordutil.logger.ExcelUtilLogger;

import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * Excel process Task <br>
 * Read data from a Excel File and convert those data into Bean List
 *
 * @author Monan
 *         created on 2018/7/27 11:31
 */
public class Excel2BeanTask extends Excel2DataBaseTask<List<Object>> {
    /**
     * Excel Sheet Iterator
     */
    private Iterator<Sheet> sheetIterator = null;

    /**
     * create a new Excel process Task
     *
     * @param document XML Document with Task Configuration
     * @param targets  A String array which storing package locations of bean or full reference of a bean
     * @param workBook Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Excel2BeanTask(Document document, String[] targets, Workbook workBook) throws IOException {
        super(document, targets, workBook);
    }

    /**
     * create a new Excel process Task
     *
     * @param taskName specific which task to read from configuration.
     * @param targets  A String array which storing package locations of bean or full reference of a bean
     * @param workBook Excel Sheet,Data Source
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Excel2BeanTask(String taskName, String[] targets, Workbook workBook) throws IOException {
        super(taskName, targets, workBook);
    }

    /**
     * create a new Excel process Task
     *
     * @param taskName  specific which task to read from configuration.
     * @param targets   A String array which storing package locations of bean or full reference of a bean
     * @param workBook  Excel workBook,work as Data Source.
     * @param sheetName specific which sheet name that to process,process all available sheet if given <code>Null</code>,which request task mapping in
     *                  configuration.
     * @throws IOException Any IO exception before processing the excel file.
     */
    public Excel2BeanTask(String taskName, String[] targets, Workbook workBook, String sheetName) throws IOException {
        super(taskName, targets, workBook, sheetName);
    }

    /**
     * 执行任务
     *
     * @return 储存Bean数据的二维数组
     * @throws ExcelProcessingException 当处理过程中遇到异常时,如超过错误限制,将会立即抛出ExcelProcessingException;
     * @throws Exception                读取XML Mapping中遇到异常
     */
    public ProcessResult<List<Object>> execute() throws ExcelProcessingException, Exception {
        try {
            Document document = XmlReadUtil.getMappingForExcelImportExportTask();
            //获取任务节点
            taskNode = ReadExcelMappingUtil.getTaskNode(document, taskName);
            //获取任务描述
            String taskComment = ReadExcelMappingUtil.getTaskComment(taskNode);
            //生成日志记录器
            this.excelUtilLogger = new ExcelUtilLogger(taskName, taskComment);
            this.sheetDatas = new HashMap<>();
            this.beanInfosFromXMLMap = new HashMap<>();
            this.taskHeaderRowNums = new HashMap<>();

            if (StringUtils.isNotBlank(sheetName)) {
                if (workBook2007 == null) {
                    provessSheet(workBook2003.getSheet(sheetName));
                } else {
                    provessSheet(workBook2007.getSheet(sheetName));
                }
            } else {
                if (workBook2007 == null) {
                    sheetIterator = workBook2003.sheetIterator();
                } else {
                    sheetIterator = workBook2007.sheetIterator();
                }
                while (sheetIterator.hasNext()) {
                    Sheet sheet = sheetIterator.next();
                    provessSheet(sheet);
                }
            }
            ProcessResult<List<Object>> processResult = new ProcessResult<List<Object>>(excelUtilLogger, sheetDatas);
            return processResult;
        } finally {
            workBook2007.close();
        }
    }

    /**
     * Process a Excel Sheet
     *
     * @param sheet
     * @throws ExcelProcessingException
     */
    private void provessSheet(Sheet sheet) throws ExcelProcessingException {
        String tempSheetName = sheet.getSheetName();
        //获取表头行数
        int taskHeaderRowNum = ReadExcelMappingUtil.getTaskHeaderRowNum(taskNode, sheetName);
        //解析XMLmapping文件
        List<BeanInfo> beanInfosFromXML = ReadExcelMappingUtil.getBeanInfosFromXML(taskNode, targets, excelUtilLogger, tempSheetName);
        //excel data to bean data
        ExcelTo2DArrayUtil.sheetToBeanInfoArray(sheet, taskHeaderRowNum, beanInfosFromXML, excelUtilLogger);
        List<List<Object>> lists = ArrayToBeanArrayUtil.data2DarrayTo2DBeanArray(sheet, taskHeaderRowNum, beanInfosFromXML, excelUtilLogger);

        //saving process result
        SheetData<List<Object>> sheetData = new SheetData<List<Object>>(sheet, lists);
        sheetDatas.put(tempSheetName, sheetData);
        beanInfosFromXMLMap.put(tempSheetName, beanInfosFromXML);
        taskHeaderRowNums.put(tempSheetName, taskHeaderRowNum);
    }

}
