package sot.task.ppttask;

import sot.bean.ProcessResult;
import sot.bean.exception.ExcelProcessingException;

/**
 * <内容说明>
 *
 * @author Monan
 *         created on 2018/8/30 10:45
 */
public abstract class PPTBaseTask<T> {
    private PPTBaseTask() {
    }

    public abstract ProcessResult<T> execute() throws ExcelProcessingException, Exception;

    /**
     * may the force be will you
     */
}
