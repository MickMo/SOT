package sot.util.excelprocessutil;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

/**
 * Excel复制工具
 *
 * @author Monan
 * created on 2018/7/30 15:09
 */
public final class ExcelCopyUtils {
    private ExcelCopyUtils() {
    }

    /**
     * 根据单元格类型复制内容
     *
     * @param oldCell 旧单元格
     * @param newCell 新单元格
     */
    public static void copyCell(Cell oldCell, Cell newCell) {
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setCellType(CellType.BLANK);
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                break;
        }
    }
}
