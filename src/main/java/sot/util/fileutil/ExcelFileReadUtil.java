package sot.util.fileutil;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Excel文件读取工具类
 *
 * @author Monan
 * created on 2018/7/19 12:08
 */
public final class ExcelFileReadUtil {
    /**
     * 私有化构造器
     */
    private ExcelFileReadUtil() {
    }

    /**
     * 日志
     */
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelFileReadUtil.class);

    /**
     * 读取Excel2003文件
     *
     * @param xlsFileInputStream 文件流
     * @return Excel工作簿对象
     * @throws IOException 读取异常
     */
    public static HSSFWorkbook readExcelFile2003(InputStream xlsFileInputStream) throws IOException {
        if (xlsFileInputStream == null) {
            throw new RuntimeException("Null InputStream!Could Not read Excel File!");
        }
        HSSFWorkbook workBook = new HSSFWorkbook(xlsFileInputStream);
        return workBook;
    }

    /**
     * 读取Excel2007文件
     *
     * @param xlsxFileInputStream 文件流
     * @return Excel工作簿对象
     * @throws IOException 读取异常
     */
    public static Workbook readExcelFile2007(InputStream xlsxFileInputStream) throws IOException {
        if (xlsxFileInputStream == null) {
            throw new RuntimeException("Null InputStream!Could Not read Excel File!");
        }
        Workbook workBook = new XSSFWorkbook(xlsxFileInputStream);
        return workBook;
    }


    /**
     * 从文件读取Excel2003版本及以后的xls文件
     *
     * @param xlsxFile Excel File对象
     * @return 返回工作簿第
     * @throws IOException IO错误
     */
    public static Workbook readExcelFile2007(File xlsxFile) throws IOException {
        if (!xlsxFile.exists()) {
            throw new RuntimeException("Excel File does't exist on path:'" + xlsxFile.getAbsolutePath() + "'!");
        }
        Workbook workBook = new XSSFWorkbook(new FileInputStream(xlsxFile));
        return workBook;
    }

    /**
     * 从文件读取Excel2003版本及以前的xlsx文件
     *
     * @param xlsFile Excel File对象
     * @return 返回工作簿第一个工作表
     * @throws IOException IO错误
     */
    public static Sheet readExcel2003File(File xlsFile) throws IOException {
        if (!xlsFile.exists()) {
            throw new RuntimeException("Excel File does't exist on path:'" + xlsFile.getAbsolutePath() + "'!");
        }
        HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(xlsFile));
        HSSFSheet sheet = workBook.getSheetAt(0);
        return sheet;
    }

    /**
     * 从文件读取Excel2003版本及以后的xls文件
     *
     * @param xlsxFile Excel File对象
     * @return 返回工作簿第一个工作表
     * @throws IOException IO错误
     */
    public static Sheet readExcel2007File(File xlsxFile) throws IOException {
        Sheet sheet = readExcelFile2007(xlsxFile).getSheetAt(0);
        return sheet;
    }

    /**
     * 从文件读取Excel2003文件中指定位置的工作表
     *
     * @param xlsFile Excel File对象
     * @param sheetNo 工作表序号,从0开始
     * @return 返回工作簿第一个工作表
     * @throws IOException IO错误
     */
    public static Sheet readExcel2003File(File xlsFile, int sheetNo) throws IOException {
        if (!xlsFile.exists()) {
            throw new RuntimeException("Excel File does't exist on path:'" + xlsFile.getAbsolutePath() + "'!");
        }
        HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(xlsFile));
        HSSFSheet sheet = workBook.getSheetAt(sheetNo);
        return sheet;
    }

    /**
     * 从文件读取Excel2007文件(Older)中指定位置的工作表
     *
     * @param xlsxFile Excel File对象
     * @param sheetNo  工作表序号,从0开始
     * @return 返回工作簿第一个工作表
     * @throws IOException IO错误
     */
    public static Sheet readExcel2007File(File xlsxFile, int sheetNo) throws IOException {
        if (!xlsxFile.exists()) {
            throw new RuntimeException("Excel File does't exist on path:'" + xlsxFile.getAbsolutePath() + "'!");
        }
        Workbook workBook = new XSSFWorkbook(new FileInputStream(xlsxFile));
        Sheet sheet = workBook.getSheetAt(sheetNo);
        return sheet;
    }
}
