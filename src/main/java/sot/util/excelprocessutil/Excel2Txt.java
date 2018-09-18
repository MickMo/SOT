package sot.util.excelprocessutil;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Excel 2 TXT工具类
 *
 * @author Monan
 * created on 2018/7/31 17:34
 */
public final class Excel2Txt {
    private Excel2Txt() { }
    /**
     * 制表符
     */
    private static final String TAB_CHARACTER = "\t";

    /**
     * 将工作表转换为字符串
     * @param sheet 工作表
     * @return 工作表(字符串)
     */
    public static String sheet2Txt(Sheet sheet) {
        return sheet2Txt(sheet, 0);
    }


    /**
     * 将工作表转换为字符串
     * @param sheet 工作表
     * @param rowToSkip 指定要跳过多少行
     * @return 工作表(字符串)
     */
    public static String sheet2Txt(Sheet sheet, int rowToSkip) {
        StringBuilder tempBuilder = new StringBuilder();
        outerFor: for (Row cells : sheet) {
            while (rowToSkip > 0) {
                rowToSkip--;
                continue outerFor;
            }
            for (Cell cell : cells) {
                tempBuilder.append(ConversionUtil.getCellValueAsString(cell));
                tempBuilder.append(TAB_CHARACTER);
            }
            tempBuilder.append(System.lineSeparator());
        }
        return tempBuilder.toString();
    }
}
