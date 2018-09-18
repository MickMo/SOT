package sot.bean.exception;

import sot.wordutil.logger.ExcelUtilLogger;

/**
 * Excel处理异常
 *
 * @author Monan
 * created on 2018/7/27 14:19
 */
public class ExcelProcessingException extends Throwable {


    /**
     * 创建一个Excel处理异常
     */
    private ExcelProcessingException() {

    }

    /**
     * 创建一个Excel处理异常
     *
     * @param message 异常消息
     */
    public ExcelProcessingException(String message) {
        super(message);

    }

    /**
     * 创建一个Excel处理异常
     *
     * @param message 异常消息
     * @param cause   异常原因
     */
    public ExcelProcessingException(String message, Throwable cause) {

        super(message, cause);
    }

    /**
     * 创建一个Excel处理异常
     *
     * @param cause 异常原因
     */
    public ExcelProcessingException(Throwable cause) {

        super(cause);
    }

    /**
     * 错误日志记录器
     */
    private ExcelUtilLogger excelUtilLogger;

    /**
     * 创建一个Excel处理异常
     *
     * @param excelUtilLogger New value of 错误日志记录器.
     */
    public ExcelProcessingException(ExcelUtilLogger excelUtilLogger, String message) {
        super();
        this.excelUtilLogger = excelUtilLogger;
    }

    /**
     * Gets 错误日志记录器.
     *
     * @return Value of 错误日志记录器.
     */
    public ExcelUtilLogger getExcelUtilLogger() {
        return excelUtilLogger;
    }


    /**
     * 获取异常原因
     *
     * @return 异常原因(字符串)
     */
    public String getCauseString() {
        return excelUtilLogger.toString();
    }


}
