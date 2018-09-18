package sot.wordutil.logger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sot.constant.SOTExcelConfigureConstant;
import sot.util.excelprocessutil.ConversionUtil;
import sot.util.excelprocessutil.ExcelCopyUtils;
import sot.bean.exception.ExcelProcessingException;
import sot.bean.BeanInfo;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 导出导入日志记录
 *
 * @author Monan
 * created on 2018/7/27 9:52
 */
public class ExcelUtilLogger {
    private ExcelUtilLogger() {
    }

    /**
     * 分割线
     */
    private String splitter = "-----------------splitter-----------------";

    /**
     * 错误Excle
     */
    private Workbook errorWorkBook = null;

    /**
     * 错误Sheet
     */
    private Sheet errorInfoSheet = null;

    /**
     * 记录当前行
     */
    private int rowPointer = 0;

    /**
     * 错误记录头
     */
    private String errorLogPrefix = "[ERROR]";

    /**
     * 警告记录头
     */
    private String wanrLogPrefix = "[WARN]";

    /**
     * 错误总数
     */
    private int errorCount = 0;


    /**
     * 警告总数
     */
    private int warnCount = 0;

    /**
     * 创建错误日志记录器
     *
     * @param taskName 任务名
     * @param comment  任务描述
     */
    public ExcelUtilLogger(String taskName, String comment) {
        //TODO 增加多sheet支持
        errorWorkBook = new XSSFWorkbook();
        errorInfoSheet = errorWorkBook.createSheet("ErrorInfo");

        //TaskName
//        Cell taskNameCell = createNewRow().createCell(0, CellType.STRING);
        Row row1 = createNewRow();
        Cell taskNameCell = row1.createCell(0);
        if (StringUtils.isEmpty(comment.trim())) {
            comment = "No description for this task";
        }
        String title = SOTExcelConfigureConstant.TASK_NAME + ":" + taskName + "," + comment;
        taskNameCell.setCellValue(title);

//        //UserInfo
//        Row row2 = createNewRow();
//        Cell userInfoCell = row2.createCell(0);
//        userInfoCell.setCellValue("UserInfo:" );

        //Date
        Row row3 = createNewRow();
        Cell dateCell = row3.createCell(0);
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        dateCell.setCellValue(simpleDateFormat.format(new Date()));

        //Summary
        Row row4 = createNewRow();
        row4.createCell(0).setCellValue("Task Summary");

        row4.createCell(2).setCellValue("Error Count:");
        Cell errorCell = row4.createCell(3);
        errorCell.setCellType(CellType.STRING);
        errorCell.setCellValue(errorCount);

        row4.createCell(4).setCellValue("Warn Count:");
        Cell warnCell = row4.createCell(5);
        errorCell.setCellType(CellType.STRING);
        errorCell.setCellValue(warnCount);

        //Indicating
        createNewRow().createCell(0).setCellValue("The error message start below this line...");
    }

    /**
     * 更新任务概括
     *
     * @param errorCount 错误计数
     * @param warnCount  警告计数
     */
    private void updateSummary(int errorCount, int warnCount) {
        Row summaryRow = errorInfoSheet.getRow(3);
        summaryRow.getCell(3).setCellValue(errorCount);
        summaryRow.getCell(5).setCellValue(warnCount);
    }

    /**
     * 记录Excel处理错误
     *
     * @param beanInfo    bean信息
     * @param originalRow 原始行数据
     * @param exception   异常数据
     * @throws ExcelProcessingException 异常
     */
    public void recordExcelError(BeanInfo beanInfo, Row originalRow, Exception exception) throws ExcelProcessingException {
        recordExcelException(beanInfo, originalRow, exception, true);
    }

    /**
     * 记录Excel处理信息
     *
     * @param beanInfo    bean信息
     * @param originalRow 原始行数据
     * @param exception   异常数据
     * @throws ExcelProcessingException 异常
     */
    public void recordExcelInfo(BeanInfo beanInfo, Row originalRow, Exception exception) throws ExcelProcessingException {
        recordExcelException(beanInfo, originalRow, exception, false);
    }


    /**
     * 记录Excel异常
     *
     * @param beanInfo    bean信息
     * @param originalRow 原始行数据
     * @param exception   异常数据
     * @param isAnError   true作为异常记录,false作为信息记录
     * @throws ExcelProcessingException 异常
     */
    public void recordExcelException(BeanInfo beanInfo, Row originalRow, Exception exception, boolean isAnError) throws ExcelProcessingException {
        Row newRow = createNewRow();
        //复制原始数据
        for (Cell cell : originalRow) {
            Cell newCell = newRow.createCell(cell.getColumnIndex());
            ExcelCopyUtils.copyCell(cell, newCell);
        }


        if (isAnError) {
            createNewRow().createCell(0).setCellValue(errorLogPrefix + " " + exception.getMessage());
            errorCount++;
            addErrorCount(beanInfo);
        } else {
            createNewRow().createCell(0).setCellValue(wanrLogPrefix + " " + exception.getMessage());
            warnCount++;
        }
        createNewRow().createCell(0).setCellValue(splitter);
        createNewRow();
        updateSummary(errorCount, warnCount);
    }

    /**
     * 记录Excel错误
     *
     * @param beanInfo bean信息
     * @throws ExcelProcessingException 异常
     */
    public void addErrorCount(BeanInfo beanInfo) throws ExcelProcessingException {
        beanInfo.addError();
        if (beanInfo.isReachErrorLimit()) {
            //到达错误限制,抛出异常
            updateSummary(errorCount, warnCount);
            createNewRow().createCell(0).setCellValue(errorLogPrefix + " Process end cause by 'Reach Max Error Limit' of " + beanInfo.getClassName());
            throw new ExcelProcessingException(this, "Reach Max Error Limit");
        }
    }


    /**
     * 记录XML解析错误
     *
     * @param exception 异常信息
     * @throws ExcelProcessingException 立即抛出具体的异常
     */
    public void recordXMLError(Exception exception) throws ExcelProcessingException {
        String localizedMessage = exception.getLocalizedMessage();
        System.out.println(localizedMessage);
        System.out.println(exception.getMessage());
        exception.printStackTrace();


        createNewRow().createCell(0).setCellValue(exception.getMessage());
        createNewRow().createCell(0).setCellValue("Causes: " + exception.getCause());
        throw new ExcelProcessingException(this, exception.getMessage());
    }


    /**
     * 创建下一个新行
     *
     * @return Row
     */
    private Row createNewRow() {
        return errorInfoSheet.createRow(rowPointer++);
    }


    @Override
    public String toString() {
        StringBuilder stringBuilder = new StringBuilder();
        for (Sheet sheet : errorWorkBook) {
            for (Row cells : sheet) {
                for (Cell cell : cells) {
                    stringBuilder.append(ConversionUtil.getCellValueAsString(cell));
                    stringBuilder.append("    ");
                }
                stringBuilder.append(System.lineSeparator());
            }
        }
        return stringBuilder.toString();
    }


    /**
     * Gets 错误Excle.
     *
     * @return Value of 错误Excle.
     */
    public Workbook getErrorWorkBook() {
        return errorWorkBook;
    }

    /**
     * Gets 错误Sheet.
     *
     * @return Value of 错误Sheet.
     */
    public Sheet getErrorInfoSheet() {
        return errorInfoSheet;
    }
}
