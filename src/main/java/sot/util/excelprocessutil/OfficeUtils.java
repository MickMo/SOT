//package sot.util.excelprocessutil;
//
//import org.apache.commons.lang3.StringUtils;
//import org.apache.poi.hwpf.HWPFDocument;
//import org.apache.poi.hwpf.usermodel.Range;
//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.util.Units;
//import org.apache.poi.xwpf.usermodel.*;
//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
//import sot.bean.OfficeReplaceTag;
//
//import javax.imageio.ImageIO;
//import java.awt.image.BufferedImage;
//import java.io.*;
//import java.util.*;
//import java.util.Map.Entry;
//
///**
// * 包含:
// * doc文件内容替换工具phasingWord.
// * 默认前缀为:'$/'
// *
// * @author Mick
// */
//public class OfficeUtils {
//
//    /**
//     * 日志
//     */
//    private static final Logger LOGGER = LoggerFactory.getLogger(OfficeUtils.class);
//
//    public static final String IMG_CONTENT = "content";
//    private static final String IMG_WIDTH = "width";
//    private static final String IMG_HEIGHT = "height";
//
//    /**
//     * 根据关键字替换Word文档中的内容
//     *
//     * @param data             数据
//     * @param document         Word 2003 工作簿对象
//     * @param officeReplaceTag 替换标签
//     * @throws IOException
//     * @throws InvalidFormatException
//     */
//    public static void generateWord(Map<String, Object> data, XWPFDocument document, OfficeReplaceTag officeReplaceTag) throws IOException,
//            InvalidFormatException {
//        LOGGER.info("开始替换处理 MS Word文件");
//
//        if (data == null || data.size() < 1) {
//            throw new RuntimeException("Empty Data Map!");
//        }
//        settingTag(officeReplaceTag);
//        // 处理段落
//        LOGGER.debug("开始处理Word文件段落...");
//        List<XWPFParagraph> paragraphList = ((XWPFDocument) document).getParagraphs();
//        processParagraphsDoc07(paragraphList, data, (XWPFDocument) document, null);
//        // 处理表格
//
//        LOGGER.debug("开始处理Word文件表格...");
//        Iterator<XWPFTable> it = document.getTablesIterator();
//        while (it.hasNext()) {
//            XWPFTable table = it.next();
//            List<XWPFTableRow> rows = table.getRows();
//            for (XWPFTableRow row : rows) {
//                List<XWPFTableCell> cells = row.getTableCells();
//                for (XWPFTableCell cell : cells) {
//                    //获取单元格尺寸
//                    int cellHeight = row.getHeight() / 15;
//                    int cellWidth = cell.getCTTc().getTcPr().getTcW().getW().intValue() / 15;
//                    Map<String, Integer> imgMap = new HashMap<String, Integer>();
//                    imgMap.put(IMG_HEIGHT, cellHeight);
//                    imgMap.put(IMG_WIDTH, cellWidth);
//                    List<XWPFParagraph> paragraphListTable = cell.getParagraphs();
//                    processParagraphsDoc07(paragraphListTable, data, (XWPFDocument) document, imgMap);
//                }
//            }
//        }
//
//    }
//
//
//    /**
//     * 根据关键字替换Word文档中的内容
//     *
//     * @param data             数据
//     * @param document         Word 2003 工作簿对象
//     * @param officeReplaceTag 替换标签
//     */
//    public static void fillingWord(Map<String, Object> data, HWPFDocument document, OfficeReplaceTag officeReplaceTag) {
//        LOGGER.info("开始替换处理 MS Word文件");
//        if (data == null || data.size() < 1) {
//            throw new RuntimeException("Empty Data Map!");
//        }
//        settingTag(officeReplaceTag);
//        replaceTextContent((HWPFDocument) document, data);
//    }
//
//    /**
//     * 根据关键字替换Word文档中的内容
//     *
//     * @param data             数据
//     * @param workbook         Word 2007 工作簿对象
//     * @param officeReplaceTag 替换标签
//     */
//    public static void fillingExcel(Map<String, Object> data, Workbook workbook, OfficeReplaceTag officeReplaceTag) throws IllegalArgumentException {
//        LOGGER.info("开始替换处理 MS Excel文件");
//        if (data == null || data.size() < 1) {
//            throw new RuntimeException("Empty Data Map!");
//        }
//        //检查是否指定了前后缀,若没有,则使用默认前后缀
//        settingTag(officeReplaceTag);
//        processingWorkbook(workbook, data);
//
//    }
//
//
//    /**
//     * 处理Excel03文件入口
//     *
//     * @param workbook 模板文件
//     * @param data     替换数据
//     */
//    private static void processingWorkbook(Workbook workbook, Map<String, Object> data) {
//        try {
//            // 遍历工作簿
//            if (workbook != null) {
//                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
//                    // 遍历工作表
//                    Sheet tempSheet = workbook.getSheetAt(i);
//                    LOGGER.info("处理表格:'" + tempSheet.getSheetName() + "'");
//                    if (tempSheet != null) {
//                        // 遍历行
//                        Iterator<Row> rowIterator = tempSheet.iterator();
//                        while (rowIterator.hasNext()) {
//                            Row tempRow = rowIterator.next();
//                            // 处理每行中的单元格
//                            processingCell(tempRow, data, tempSheet, workbook);
//                        }
//                    } else {
//                        //空表
//                        LOGGER.info("空表格,跳过");
//                    }
//                }
//            } else {
//                //空工作簿
//                LOGGER.info("空工作簿,跳过");
//            }
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//
//
//    /**
//     * 适用于Excel
//     * 替换处理表格中匹配的内容
//     *
//     * @param tempRow  当前行
//     * @param data     替换数据
//     * @param sheet    当前表
//     * @param workbook 模板
//     * @throws InvalidFormatException 图片数据非法或找不到图片文件
//     * @throws IOException            IO错误
//     */
//    private static void processingCell(Row tempRow, Map<String, Object> data, Sheet sheet, Workbook workbook) throws InvalidFormatException, IOException {
//        //遍历单元格
//        if (tempRow != null) {
//            Iterator<Cell> cellIterator = tempRow.iterator();
//            //noinspection WhileLoopReplaceableByForEach
//            while (cellIterator.hasNext()) {
//                Cell tempCell = cellIterator.next();
//                //单元格格式为文本,进行处理
//                if (tempCell.getCellType() != CellType.FORMULA && tempCell.getCellType() != CellType.BLANK) {
//                    int intColumnIndex = tempCell.getColumnIndex();
//                    int intRowIndex = tempCell.getRowIndex();
//
//                    int intColumnIndex65 = intColumnIndex + 65;
//                    char charColumnIndex = (char) intColumnIndex65;
//                    LOGGER.debug("目标单元格'" + charColumnIndex + "-" + (intRowIndex + 1) + "'");
//
//                    String stringCellValue = tempCell.getStringCellValue();//文本框内容
//                    Set<Entry<String, Object>> entrySet = data.entrySet();
//                    for (Entry<String, Object> entry : entrySet) {
//                        String targetString = (preFix + entry.getKey()).toUpperCase();//大写(替换标签前缀+替换标签文本)
//                        Object value = entry.getValue();
//                        if ((stringCellValue.toUpperCase()).contains(targetString)) {
//                            //文本替换
//                            if (value instanceof String) {
//                                String strValue = value.toString();
//
//                                LOGGER.debug("找到文字替换标签");
//                                LOGGER.debug("当前内容:'" + stringCellValue + "'" + "  ,替换目标:'" + targetString + "'" + "  ,替换为:'" + strValue + "'");
//
//                                //替换处理
//                                tempCell.setCellValue(stringCellValue.replace(targetString, strValue));
//                            }
//                            //图片替换
//                            else if (value instanceof Map) {
//                                //获取图片格式参数
//                                LOGGER.debug("找到图片替换标签" + targetString);
//
//                                @SuppressWarnings("unchecked")
//                                Map<String, Object> pic = (Map<String, Object>) value;
//
//                                //图片数据合法,进行图片替换处理
//                                if (pic.get(IMG_CONTENT) != null) {
//
//                                    //对多图片插入的一个保留
//                                    //保留map中图片格式,如果有的话
//                                    //noinspection Duplicates
//                                    boolean hasOriginalImgParam = false;//默认图片Map数据中没有合法的尺寸
//                                    Object originalWidth = pic.get(IMG_WIDTH);
//                                    Object originalHeight = pic.get(IMG_HEIGHT);
//                                    if (originalWidth != null && originalHeight != null) {
//                                        //图片Map数据中合法的图片尺寸
//                                        hasOriginalImgParam = true;
//                                    }
//
//                                    //尝试获取图片属相
//                                    //尝试读取图片格式标签
//                                    int[] imgSize = analysImgPramaTag(stringCellValue);
//                                    //校验标签数据
//                                    //noinspection Duplicates
//                                    if (imgSize == null) {
//                                        imgSize = new int[(2)];
//                                        //标签校验失败,宽度或高度之一数据非法
//                                        //noinspection Duplicates
//                                        if (hasOriginalImgParam) {
//                                            //若图片Map中存在合法数据,尝试在图片Map中查找参数
//                                            LOGGER.debug("Debug :尝试在替换数据Map中查找参数");
//
//                                            imgSize[0] = Integer.parseInt(originalWidth.toString());
//                                            imgSize[1] = Integer.parseInt(originalHeight.toString());
//                                        } else {
//                                            //没有合法数据,给默认值0,程序后面后自动获取图片原始大小并按比例缩放
//                                            imgSize[0] = 0;
//                                            imgSize[1] = 0;
//                                            LOGGER.debug("没有找到图片参数,将使用图片原始数据");
//                                        }
//                                    } else {
//                                        //图片格式标签解析完成
//                                        LOGGER.debug("图片格式标签解析结果W*H  :  " + imgSize[0] + " * " + imgSize[1]);
//                                    }
//
//                                    //将参数放入参数Map
//                                    pic.put(IMG_WIDTH, imgSize[0]);
//                                    pic.put(IMG_HEIGHT, imgSize[1]);
//
//                                    //开始替换处理
//                                    //1.移除所有文字
//                                    tempCell.setCellValue("");
//                                    //2.添加图片
//                                    replaceImgInExcel(pic, tempCell, intColumnIndex, intRowIndex, tempRow, sheet, workbook);
//
//                                    //恢复原始图片格式,如果有的话
//                                    //noinspection Duplicates
//                                    if (hasOriginalImgParam) {
//                                        pic.put(IMG_WIDTH, originalWidth);
//                                        pic.put(IMG_HEIGHT, originalHeight);
//                                    } else {
//                                        //原始Map中不存在图片格式,尝试移除
//                                        pic.remove(IMG_WIDTH);
//                                        pic.remove(IMG_HEIGHT);
//                                    }
//                                } else {
//                                    //图片数据非法,不处理
//                                    LOGGER.error("图片数据非法:没有找到图片数据,不做处理");
//                                }
//                            } else {
//                                //暂不支持的替换内容
//                                LOGGER.error("不支持的替换--替换标签:" + targetString + "  替换数据类型:" + value.getClass() + ",不做处理");
//                            }
//                        }
//                    }
//                } else {
//                    //单元格格式不为文本,跳过
//                    int intColumnIndex = tempCell.getColumnIndex() + 65;
//                    char charColumnIndex = (char) intColumnIndex;
//                    int intRowIndex = tempCell.getRowIndex() + 1;
//                    LOGGER.error("当前位置:" + charColumnIndex + "-" + intRowIndex + "  不支持的单元格格式:'" + tempCell.getCellStyle() + "'  当前只支持文本类型");
//
//                }
//            }
//        } else {
//            //行为空
//            LOGGER.debug("表格行为空,跳过");
//        }
//    }
//
//    /**
//     * 适用于Word2003
//     * 替换处理入口
//     * 获取段落内容,替换掉所有匹配的替换标签,然后更新段落内容
//     *
//     * @param doc  当前文档
//     * @param data 替换数据Map
//     */
//    private static void replaceTextContent(HWPFDocument doc, Map<String, Object> data) {
//        try {
//            Range range = doc.getRange();
//            Set<Entry<String, Object>> entrySet = data.entrySet();
//            for (Entry<String, Object> entry : entrySet) {
//                Object value = entry.getValue();
//                //替换文字
//                if (value instanceof String) {
//                    String strValue = value.toString();
//                    if (StringUtils.isNotBlank(strValue)) {
//                        range.replaceText(preFix + entry.getKey().toUpperCase(), strValue);
//                    }
//                }
//                //替换图片
//                else if (value instanceof Map) {
//                    LOGGER.error("POI不支持在03版本的Word文件中插入图片");
//                }
//                //暂不支持的其他类型
//                else {
//                    LOGGER.error("不支持的替换数据类型:" + value.getClass());
//                }
//            }
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//
//
//    /**
//     * 适用于Word2007+
//     * 替换处理入口
//     * 替换段落Paragraph中匹配的内容
//     *
//     * @param paragraphList 段落List
//     * @param data          替换数据Map
//     * @param doc           当前文档
//     * @throws IOException            替换图片路径未找到.
//     * @throws InvalidFormatException 图片文件错误或图片文件路径不合法
//     */
//    private static void processParagraphsDoc07(List<XWPFParagraph> paragraphList, Map<String, Object> data, XWPFDocument doc, Map<String, Integer>
//            imgMap)
//            throws IOException, InvalidFormatException {
//        if (paragraphList != null && paragraphList.size() > 0) {
//
//            for (XWPFParagraph paragraph : paragraphList) {
//                String paragraphText = paragraph.getText().toUpperCase();
//                LOGGER.debug("当前段落文本:'" + paragraphText + "'");
//                if (StringUtils.isNotBlank(paragraphText)) {
//                    // 遍历参数Map,寻找可匹配对象
//                    for (Entry<String, Object> entry : data.entrySet()) {
//                        String key = preFix + entry.getKey().toUpperCase();
//                        String targetText = entry.getKey().toUpperCase();
//                        Object value = entry.getValue();
//                        //找到替换目标
//                        if (paragraphText.contains(key)) {
//                            // 文本替换处理
//                            if (value instanceof String) {
//
//                                LOGGER.debug("段落中找到替换标签:'" + key + "'");
//
//                                //查找替换关键字所在的run
//                                List<XWPFRun> runs = paragraph.getRuns();
//                                int runNum = runs.size();
//                                String strValue = value.toString();
//                                boolean isPreRunContainPreFix = false;//用于记录前一个Run是否匹配了标签前缀
//
//                                LOGGER.debug("开始遍历run...");
//
//                                for (int i = 0; i < runNum; i++) {
//                                    String tempRunText = runs.get(i).getText(0).toUpperCase();
//
//                                    if (tempRunText.contains(preFix) && !isPreRunContainPreFix) {
//                                        //当前段落包含替换标签前缀
//                                        LOGGER.debug(" -> 找到Prefix:'" + preFix + "'");
//
//                                        if (tempRunText.contains(key)) {
//                                            //进一步判断:包含替换标签+替换文本
//                                            LOGGER.debug("找到完整替换标签,开始替换处理...");
//
//                                            //找到目标,进行替换
//                                            runs.get(i).setText(strValue, 0);
//                                            //替换完成,继续在同一段落中检查替换
//                                        } else {
//                                            //仅包含替换标签,下一循环开始匹配替换标签中的文本部分
//                                            isPreRunContainPreFix = true;
//                                            LOGGER.debug("在下一run中查找替换标签正文");
//                                        }
//                                    } else //noinspection ConstantConditions
//                                        if (isPreRunContainPreFix) {
//
//                                            //仅在上一个run中找到替换标签时才会检查当前run中是否存在替换标签文本部分
//                                            //包含替换标签文本部分
//                                            if (tempRunText.contains(targetText)) {
//                                                LOGGER.debug("找到替换标签正文:'" + targetText + "',开始替换处理...");
//
//                                                //进行替换
//                                                runs.get(i).setText(strValue, 0);
//                                                //移除上一个run中的替换标签
//                                                runs.get(i - 1).setText("", 0);
//                                                //替换完成,继续在同一段落中检查替换
//                                            } else if (tempRunText.contains(preFix)) {
//                                                //移除上一个run中的替换标签
//                                                runs.get(i - 1).setText("", 0);
//                                                LOGGER.debug("找到Prefix:'" + preFix + "',删除上一标签前缀,在下一run中查找替换标签正文...");
//
//                                                //继续匹配替换标签的文本部分
//                                            } else {
//                                                //仅匹配到标签前缀
//                                                //移除标签
//                                                runs.get(i - 1).setText("", 0);
//                                                LOGGER.debug("找不到替换标签正文,删除该标签前缀");
//
//                                                //开始新的匹配
//                                                isPreRunContainPreFix = false;
//                                            }
//                                        }
//                                }
//                            }
//
//                            // 图片替换
//                            else if (value instanceof Map) {
//                                LOGGER.debug("找到图片替换标签" + key);
//
//                                @SuppressWarnings("unchecked")
//                                Map<String, Object> pic = (Map<String, Object>) value;
//
//                                //判断图片文件地址是否存在
//                                if (pic.get(IMG_CONTENT) != null) {
//                                    //对多图片插入的一个保留
//                                    //保留map中图片格式,如果有的话
//                                    //noinspection Duplicates
//                                    boolean hasOriginalImgParam = false;//图片Map中是否存在合法数据
//                                    Object originalWidth = pic.get(IMG_WIDTH);
//                                    Object originalHeight = pic.get(IMG_HEIGHT);
//                                    //TODO 整个数据校验代码需要提取成一个方法,并在替换操作前执行
//                                    if (StringUtils.isNotBlank(originalWidth.toString()) && StringUtils.isNotBlank(originalHeight.toString())) {
//                                        //图片Map数据中合法的图片尺寸
//                                        hasOriginalImgParam = true;
//                                    }
//
//                                    //尝试获取图片属相
//                                    //尝试读取图片格式标签
//                                    int[] imgSize = analysImgPramaTag(paragraphText);
//                                    //校验标签数据
//                                    //noinspection Duplicates
//                                    if (imgSize == null) {
//                                        imgSize = new int[(2)];
//                                        //标签校验失败,宽度或高度之一数据非法
//                                        //noinspection Duplicates
//                                        if (hasOriginalImgParam) {
//                                            //若图片Map中存在合法数据,尝试在图片Map中查找参数
//                                            LOGGER.debug("尝试在替换数据Map中查找图片参数");
//
//                                            imgSize[0] = Integer.parseInt(originalWidth.toString());
//                                            imgSize[1] = Integer.parseInt(originalHeight.toString());
//                                        } else {
//
//
//                                            if (imgMap == null) {
//                                                //没有合法数据,给默认值0,程序后面后自动获取图片原始大小并按比例缩放
//                                                imgSize[0] = 0;
//                                                imgSize[1] = 0;
//                                                LOGGER.debug("没有找到图片参数,将使用图片原始数据");
//                                            }
//
//                                            //使用单元格数据
//                                            else {
//                                                imgSize[0] = imgMap.get(IMG_WIDTH);
//                                                imgSize[1] = imgMap.get(IMG_HEIGHT);
//                                                LOGGER.debug("使用单元格大小数据:单元高:" + imgSize[1] + ",   单元宽:" + imgSize[0]);
//                                            }
//
//
//                                        }
//                                    } else {
//                                        //图片格式标签解析完毕
//                                        LOGGER.debug("图片格式标签解析结果W*H  :" + imgSize[0] + " * " + imgSize[1]);
//
//                                    }
//
//                                    //将参数放入参数Map
//                                    pic.put(IMG_WIDTH, imgSize[0]);
//                                    pic.put(IMG_HEIGHT, imgSize[1]);
//                                    //开始替换处理
//                                    //1.移除所有Run
//                                    List<XWPFRun> runs = paragraph.getRuns();
//                                    for (XWPFRun xwpfRun : runs) {
//                                        xwpfRun.setText("", 0);
//                                    }
//                                    for (int i = 0; i < runs.size() + 1; i++) {
//                                        paragraph.removeRun(0);
//                                    }
//
//                                    LOGGER.debug("清除段落文本内容,并插入图片");
//
//                                    XWPFRun imgRun = paragraph.createRun();
//                                    //2.添加图片
//                                    replaceImginWord(pic, imgRun, doc);
//
//                                    //恢复原始图片格式,如果有的话
//                                    //noinspection Duplicates
//                                    if (hasOriginalImgParam) {
//                                        pic.put(IMG_WIDTH, originalWidth);
//                                        pic.put(IMG_HEIGHT, originalHeight);
//                                    }
//                                    //原始Map中不存在图片格式,尝试移除
//                                    else {
//                                        pic.remove(IMG_WIDTH);
//                                        pic.remove(IMG_HEIGHT);
//                                    }
//
//                                    //完成图片替换
//
//                                } else {
//                                    //图片数据非法
//                                    LOGGER.debug("图片数据非法:没有找到图片数据");
//
//                                }
//                                //继续匹配下一关键字替换
//                            }
//                        }
//                    }
//                    LOGGER.debug("debug: " + "Debug :当前行处理完毕");
//
//                } else {
//                    // 空行
//                    LOGGER.debug("空行,不处理");
//
//                }
//            }
//        }
//    }
//
//
//    /**
//     * 处理图片格式标签
//     *
//     * @param paragraphText 含有标签的文本
//     * @return 读取到的图片格式:W*H
//     */
//    private static int[] analysImgPramaTag(String paragraphText) {
//        int[] imgSize = null;
//        //尝试读取图片格式标签
//        //默认的图片格式标签<宽:100,高:100>
//        if (paragraphText.contains(imgPreFix) && paragraphText.contains(imgSuffix)) {
//            //宽:100,高:100
//            String imgParam = paragraphText.substring(paragraphText.lastIndexOf(imgPreFix) + 1, paragraphText.lastIndexOf(imgSuffix));
//            LOGGER.debug("开始处理图片格式标签:" + imgParam);
//
//            //校验标签
//            if (imgParam.contains("宽") && imgParam.contains("高") && imgParam.contains(":")) {
//                //[宽:100],[高:100]
//                String[] imgParamSplit = imgParam.split(",");
//                imgSize = new int[(2)];
//                for (String paramBlock : imgParamSplit) {
//                    //[宽],[100]
//                    String[] paramElement = paramBlock.split(":");
//                    if (paramElement[0].equals("宽")) {
//                        imgSize[0] = Integer.parseInt(paramElement[1]);
//                    } else if (paramElement[0].equals("高")) {
//                        imgSize[1] = Integer.parseInt(paramElement[1]);
//                    }
//                }
//            } else {
//                //找到的标签不合法
//                LOGGER.debug("图片格式标签不合法.");
//
//            }
//        } else {
//            LOGGER.debug("图片格式标签前后缀不匹配,当前图片格式标签为  前缀:" + imgPreFix + "  后缀:" + imgSuffix);
//
//        }
//        return imgSize;
//    }
//
//
//    /**
//     * 适用于Word
//     * 根据图片类型，取得对应的图片类型代码
//     *
//     * @param picType 图片后缀名
//     * @return int 图片类型码
//     */
//    private static int getPictureTypeInWord(String picType) {
//        int res = XWPFDocument.PICTURE_TYPE_PICT;
//        if (picType != null) {
//            if (picType.equalsIgnoreCase("png")) {
//                res = XWPFDocument.PICTURE_TYPE_PNG;
//            } else if (picType.equalsIgnoreCase("dib")) {
//                res = XWPFDocument.PICTURE_TYPE_DIB;
//            } else if (picType.equalsIgnoreCase("emf")) {
//                res = XWPFDocument.PICTURE_TYPE_EMF;
//            } else if (picType.equalsIgnoreCase("jpg")
//                    || picType.equalsIgnoreCase("jpeg")) {
//                res = XWPFDocument.PICTURE_TYPE_JPEG;
//            } else if (picType.equalsIgnoreCase("wmf")) {
//                res = XWPFDocument.PICTURE_TYPE_WMF;
//            }
//        }
//        return res;
//    }
//
//
//    /**
//     * 适用于Excel
//     * 根据图片类型，取得对应的图片类型代码
//     *
//     * @param picType 图片后缀名
//     * @return int 图片类型码
//     */
//    private static int getPictureTypeinExcel(String picType) {
//
//        int res = Workbook.PICTURE_TYPE_PICT;
//        if (picType != null) {
//            if (picType.equalsIgnoreCase("png")) {
//                res = Workbook.PICTURE_TYPE_PNG;
//            } else if (picType.equalsIgnoreCase("dib")) {
//                res = Workbook.PICTURE_TYPE_DIB;
//            } else if (picType.equalsIgnoreCase("emf")) {
//                res = Workbook.PICTURE_TYPE_EMF;
//            } else if (picType.equalsIgnoreCase("jpg")
//                    || picType.equalsIgnoreCase("jpeg")) {
//                res = Workbook.PICTURE_TYPE_JPEG;
//            } else if (picType.equalsIgnoreCase("wmf")) {
//                res = Workbook.PICTURE_TYPE_WMF;
//            }
//        }
//        return res;
//    }
//
//    private static void settingTag(OfficeReplaceTag officeReplaceTag) {
//        //检查是否指定了前后缀,若没有,则使用默认前后缀
//        if (officeReplaceTag != null) {
//            String tpreFix = officeReplaceTag.getPreFix();
//            String timgPreFix = officeReplaceTag.getImgPreFix();
//            String timgSuffix = officeReplaceTag.getImgSuffix();
//            if (StringUtils.isNotBlank(tpreFix)) {
//                preFix = tpreFix;
//            }
//            if (StringUtils.isNotBlank(timgPreFix)) {
//                imgPreFix = timgPreFix;
//            }
//            if (StringUtils.isNotBlank(timgSuffix)) {
//                imgSuffix = timgSuffix;
//            }
//
//            LOGGER.info("使用指定标签前缀:");
//            LOGGER.info("preFix:" + preFix);
//            LOGGER.info("imgPreFix:" + imgPreFix);
//            LOGGER.info("imgPreFix:" + imgPreFix);
//        } else {
//            LOGGER.info("使用默认标签前缀:");
//            LOGGER.info("preFix:" + preFix);
//            LOGGER.info("imgPreFix:" + imgPreFix);
//            LOGGER.info("imgPreFix:" + imgPreFix);
//        }
//    }
//
//
//    /**
//     * 将输入流中的数据写入字节数组
//     * 尽量不要用变量接受,用完即GC
//     *
//     * @param in 输入流
//     * @return 字节数组
//     */
//    private static byte[] inputStream2ByteArray(InputStream in) throws IOException {
//        byte[] byteArray;
//        try {
//            int total = in.available();
//            byteArray = new byte[total];
//            //noinspection ResultOfMethodCallIgnored
//            in.read(byteArray);
//        } catch (IOException e) {
//            e.printStackTrace();
//            throw e;
//        }
//        return byteArray;
//    }
//
//
//    /**
//     * 自动按比例缩放到目标尺寸
//     * 不是拉伸图片
//     *
//     * @param targetSize      目标尺寸 W*H
//     * @param fileInputStream 图片数据输入流
//     * @return 缩放后的尺寸
//     * @throws IOException 无法读取图片元数据
//     */
//    private static int[] calculateSize(int[] targetSize, FileInputStream fileInputStream) throws IOException {
//
//        if (targetSize == null || targetSize.length < 1) {
//            LOGGER.error("目标尺寸非法或为空");
//            return null;
//        } else {
//            LOGGER.debug("开始缩放处理,目标尺寸 W*H :" + targetSize[0] + " * " + targetSize[1]);
//
//            BufferedImage imgSource = ImageIO.read(fileInputStream);
//            int[] imgSize = new int[(2)];
//            imgSize[0] = imgSource.getWidth();
//            imgSize[1] = imgSource.getHeight();
//            //noinspection UnusedAssignment
//            imgSource = null;//手动释放资源
//            LOGGER.debug("图片原始尺寸 W*H :  " + imgSize[0] + " * " + imgSize[1]);
//
//            int cellWidth = targetSize[0];
//            int cellHeight = targetSize[1];
//            //用于记录单元格哪一边最短,0为W,1为H
//            int maxCellLengthBorder = cellHeight > cellWidth ? 0 : 1;
//            //缩放比例
//            float resizeScale;
//            resizeScale = (float) targetSize[maxCellLengthBorder] / (float) imgSize[maxCellLengthBorder];
//            LOGGER.debug("图片缩放比例为:" + resizeScale);
//
//            //缩放
//            imgSize[0] = (int) (imgSize[0] * resizeScale);
//            imgSize[1] = (int) (imgSize[1] * resizeScale);
//            return imgSize;
//        }
//    }
//
//
//    /**
//     * 向Word中插入数据
//     * 适用于Word07
//     *
//     * @param map 图片数据
//     * @param run 当前Run
//     * @param doc 当前文档
//     * @throws IOException            流错误,图片路径不合法
//     * @throws InvalidFormatException 获取单元格尺寸失败
//     */
//    private static void replaceImginWord(Map<String, Object> map, XWPFRun run, XWPFDocument doc) throws IOException, InvalidFormatException {
//
//        Object objImgFilePath = map.get(IMG_CONTENT);
//        if (objImgFilePath != null) {
//            String imgFilePath = (String) objImgFilePath;
//
//            //先读取图片数据,避免后面ImageIO读取数据后关闭流而导致图片数据为空
//            FileInputStream fileInputStream = new FileInputStream(imgFilePath);
//            ByteArrayInputStream byteInputStream = null;
//            try {
//                //图片属性
//                String imgFile = new File(imgFilePath).getName();
//                int width = Integer.parseInt(map.get(IMG_WIDTH).toString());
//                int height = Integer.parseInt(map.get(IMG_HEIGHT).toString());
//
//                //判断图片数据Map是否合法
//                if (width != 0 && height != 0) {
//                    //有用户指定尺寸,且合法,进行自动缩放
//                    int[] targetImgSize = new int[(2)];
//                    targetImgSize[0] = width;
//                    targetImgSize[1] = height;
//                    //按比例缩放,包含ImageIO
//                    targetImgSize = calculateSize(targetImgSize, fileInputStream);
//                    if (targetImgSize != null) {
//                        width = targetImgSize[0];
//                        height = targetImgSize[1];
//                    } else {
//                        //无法获取合法的单元格尺寸
//                        LOGGER.error("替换数据中图片尺寸数据非法 或 无法获取图片原始尺寸 或 无法获取单元格尺寸");
//                        throw new InvalidFormatException("替换数据中图片尺寸数据非法 或 无法获取图片原始尺寸 或 无法获取单元格尺寸");
//                    }
//                } else {
//                    //用户没有指定图片尺寸,使用图片元数据里的尺寸数据
//                    BufferedImage imgSource = ImageIO.read(fileInputStream);
//                    width = imgSource.getWidth();
//                    height = imgSource.getHeight();
//                    //noinspection UnusedAssignment
//                    imgSource = null;//手动释放
//                }
//
//                //处理图片属性
//                String[] tempType = imgFile.split(",");
//                String imgType = tempType[tempType.length - 1];
//                int picType = getPictureTypeInWord(imgType);
//
//                LOGGER.debug("开始插入图片...");
//                LOGGER.debug("  图片地址:" + imgFilePath);
//                LOGGER.debug("  图片名:" + imgFile);
//                LOGGER.debug("  图片类型:" + picType);
//                LOGGER.debug("  插入大小W*H: " + width + "X" + height);
//
//
//                //替换图片
//                fileInputStream = new FileInputStream(imgFilePath);
//                byteInputStream = new ByteArrayInputStream(inputStream2ByteArray(fileInputStream));
//
//                //POI3.16中的API
//                run.addPicture(fileInputStream, picType, imgFile, Units.toEMU(width), Units.toEMU(height));
//
//            } catch (IOException e) {
//                e.printStackTrace();
//                throw e;
//            } finally {
//                //noinspection ConstantConditions
//                if (fileInputStream != null) {
//                    try {
//                        fileInputStream.close();
//                    } catch (IOException e) {
//                        e.printStackTrace();
//                    }
//                }
//                if (byteInputStream != null) {
//                    try {
//                        fileInputStream.close();
//                    } catch (IOException e) {
//                        e.printStackTrace();
//                    }
//                }
//            }
//        } else {
//            //图片数据非法
//            LOGGER.error("图片数据非法 : 图片路径为空");
//            throw new InvalidFormatException("图片数据非法 : 图片路径为空");
//        }
//    }
//
//
//    /**
//     * 向Excel中插入图片数据
//     *
//     * @param imgDataMap     图片数据
//     * @param targetCell     目标单元格
//     * @param intColumnIndex 当前行号
//     * @param intRowIndex    当前列号
//     * @param tempRow        当前列
//     * @param sheet          当前表
//     * @param workbook       当前工作簿
//     * @throws IOException            流错误,图片路径不合法
//     * @throws InvalidFormatException 获取单元格尺寸失败
//     */
//    private static void replaceImgInExcel(Map<String, Object> imgDataMap, Cell targetCell, int intColumnIndex, int intRowIndex, Row tempRow, Sheet sheet,
//                                          Workbook
//                                                  workbook) throws IOException, InvalidFormatException {
//
//        Object objImgFilePath = imgDataMap.get(IMG_CONTENT);
//        if (objImgFilePath != null) {
//            String imgFilePath = (String) objImgFilePath;
//            FileInputStream fileInputStream = new FileInputStream(imgFilePath);
//            try {
//                //图片属性
//                String imgFile = imgFilePath.substring(imgFilePath.lastIndexOf("\\"), imgFilePath.length());
//                String[] tempType = imgFile.split(",");
//                String imgType = tempType[tempType.length - 1];
//                int picType = getPictureTypeinExcel(imgType);
//
//                //设置单元格格式
//                LOGGER.debug("开始设置图片格式...");
//                CellStyle cellStyle = targetCell.getCellStyle();
//                cellStyle.setAlignment(HorizontalAlignment.CENTER);
//                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//
//                //获取单元格尺寸
//                int cellHeight = tempRow.getHeight() / 15;
//                int columnWidth = (int) (sheet.getColumnWidth(intColumnIndex) / 36.7);
//                int[] cellSize = new int[(2)];
//                cellSize[0] = columnWidth;
//                cellSize[1] = cellHeight;
//
//                LOGGER.debug("当前元格大小 W*H :" + columnWidth + " * " + cellHeight);
//
//                int intWidth;
//                short shortHeight;
//                Object objWidth = imgDataMap.get(IMG_WIDTH);
//                Object objHeight = imgDataMap.get(IMG_HEIGHT);
//
//                //尝试获取替换数据中图片尺寸
//                if (objHeight != null && objWidth != null && StringUtils.isNotBlank(objHeight.toString()) && StringUtils.isNotBlank(objWidth.toString())) {
//                    intWidth = Integer.parseInt(objWidth.toString());
//                    shortHeight = (short) Integer.parseInt(objHeight.toString());
//                    //检查用户指定尺寸是否超过单元格尺寸
//                    if (
//                            !(intWidth > cellSize[0]) &&
//                                    !(shortHeight > cellSize[1]) &&
//                                    intWidth > 0 &&
//                                    shortHeight > 0
//                            ) {
//
//                        //自动按比例缩放
//                        cellSize[0] = intWidth;
//                        cellSize[1] = shortHeight;
//                        int[] imgSize = calculateSize(cellSize, fileInputStream);
//                        //noinspection Duplicates
//                        if (imgSize != null) {
//                            intWidth = imgSize[0];
//                            shortHeight = (short) imgSize[1];
//                            LOGGER.debug("图片缩放后尺寸: W*H " + intWidth + " * " + shortHeight);
//
//                        } else {
//                            //无法获取合法的单元格尺寸
//                            LOGGER.error("替换数据中图片尺寸数据非法 或 无法获取图片原始尺寸 或 无法获取单元格尺寸");
//                            throw new InvalidFormatException("替换数据中图片尺寸数据非法 或 无法获取图片原始尺寸 或 无法获取单元格尺寸");
//                        }
//                    }
//                } else {
//                    // 用户未指定图片尺寸或尺寸超过单元格尺寸(会导致排版问题)
//                    // 依据单元格尺寸自动按比例缩放,并自动填充
//                    LOGGER.debug("未指定图片尺寸,将自动缩放并填充单元格");
//
//                    //自动计算缩放后的宽高
//                    int[] imgSize = calculateSize(cellSize, fileInputStream);
//                    //noinspection Duplicates
//                    if (imgSize != null) {
//                        intWidth = imgSize[0];
//                        shortHeight = (short) imgSize[1];
//                        LOGGER.debug("缩放后尺寸: W*H " + intWidth + " * " + shortHeight);
//
//                    } else {
//                        //无法获取合法的单元格尺寸
//                        LOGGER.error("无法获得图片或单元格尺寸!");
//                        throw new InvalidFormatException("无法获取合法的图片或单元格尺寸");
//                    }
//                }
//
//
//                LOGGER.debug("开始插入图片数据...");
//                LOGGER.debug("  图片地址:" + imgFilePath);
//                LOGGER.debug("  图片名:" + imgFile);
//                LOGGER.debug("  图片类型:" + picType);
//
//                //创建图片
//                CreationHelper helper = workbook.getCreationHelper();
//                Drawing drawing = sheet.createDrawingPatriarch();
//                ClientAnchor anchor = helper.createClientAnchor();
//
//                //设置图片位置
//                anchor.setCol1(intColumnIndex);
//                anchor.setRow1(intRowIndex);
//                anchor.setCol2(intColumnIndex + 1);
//                anchor.setRow2(intRowIndex + 1);
//
//                //添加图片数据
//
//                fileInputStream = new FileInputStream(imgFilePath);
//                int pictureIndex = workbook.addPicture(inputStream2ByteArray(fileInputStream), picType);
//                drawing.createPicture(anchor, pictureIndex);
//
//                //设置图片大小
//                if (intWidth != 0 && shortHeight != 0) {
//                    LOGGER.debug("设置单元格 W*H: " + intWidth + " * " + shortHeight);
//
//                    sheet.setColumnWidth(intColumnIndex, (int) (intWidth * 36.37));
//                    tempRow.setHeight((short) (shortHeight * 15));
//                }
//            } catch (IOException e) {
//                e.printStackTrace();
//                throw e;
//            } finally {
//                //noinspection ConstantConditions
//                if (fileInputStream != null) {
//                    try {
//                        fileInputStream.close();
//                    } catch (IOException e) {
//                        e.printStackTrace();
//                    }
//                }
//            }
//        } else {
//            //图片数据非法
//            LOGGER.error("图片数据非法 : 图片路径为空");
//            throw new InvalidFormatException("图片数据非法 : 图片路径为空");
//        }
//
//    }
//}