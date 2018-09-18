package sot.bean.sheetdata.impl;

import org.apache.poi.ss.usermodel.Sheet;
import sot.bean.sheetdata.BaseSheetData;

import java.util.List;

/**
 * <内容说明>
 *
 * @author Monan
 *         created on 2018/8/30 15:33
 */
public final class SheetData<T> extends BaseSheetData<T> {

    public SheetData(Sheet sheet, List<T> rowData) {
        super(sheet, rowData);
    }

    public SheetData(List<T> rowData) {
        super(rowData);
    }

}
