//package sot.util;
//
//import sot.bean.rowdata.BeanRowData;
//import sot.bean.rowdata.MapRowData;
//import sot.bean.sheetdata.impl.SheetData;
//
//import java.util.ArrayList;
//import java.util.List;
//import java.util.Map;
//
///**
// * Beans to BeanRowData Converter
// *
// * @author Monan
// *         created on 9/14/2018 9:51 PM
// */
//public final class SheetDataConverter {
//    private SheetDataConverter() {
//    }
//
//    public static SheetData<BeanRowData> convertBean(List<List> datas) {
//        List<BeanRowData> beanRowDataList = new ArrayList<>();
//        for (List data : datas) {
//            BeanRowData beanRowData = new BeanRowData(data);
//            beanRowDataList.add(beanRowData);
//        }
//        SheetData<BeanRowData> sheetData = new SheetData<BeanRowData>(beanRowDataList);
//        return sheetData;
//    }
//
//
//    public static SheetData<MapRowData> convertMap(List<Map> datas) {
//
//        List<MapRowData> mapRowDataList = new ArrayList<>();
//        for (Map data : datas) {
//            MapRowData mapRowData = new MapRowData(data);
//            mapRowDataList.add(mapRowData);
//        }
//        SheetData<MapRowData> sheetData = new SheetData<MapRowData>(mapRowDataList);
//        return sheetData;
//    }
//}
