package sot.bean.sheetdata;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

/**
 * <内容说明>
 *
 * @author Monan
 *         created on 2018/8/30 15:46
 */
public abstract class BaseSheetData<T> {
    private BaseSheetData() {
    }

    public BaseSheetData(Sheet sheet, List<T> rowData) {
        this.sheet = sheet;
        this.rowData = rowData;
    }

    public BaseSheetData(List<T> rowData) {
        this.rowData = rowData;
    }

    /**
     * Excel Sheet
     */
    private Sheet sheet;

    /**
     *
     */
    private List<T> rowData;

    /**
     *
     */
    public List<T> getRowData() {
        return rowData;
    }

    /**
     * Gets Excel Sheet.
     *
     * @return Value of Excel Sheet.
     */
    public Sheet getSheet() {
        return sheet;
    }
}
