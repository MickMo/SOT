package sot.util.excelprocessutil;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import static org.apache.poi.ss.usermodel.CellType.*;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

/**
 * <内容说明>
 *
 * @author Monan
 * created on 2018/7/19 16:43
 */
public final class ConversionUtil {
    private ConversionUtil() {

    }


    /**
     * 自动转换日期字符串
     * 日期类型我们支持如下格式  <br/>
     * yyyy-MM-dd HH:mm:ss  <br/>
     * yyyy/MM/dd <br/>
     * HH:mm:ss <br/>
     * yyyy-MM-dd <br/>
     * yyyy/MM/dd <br/>
     * yyyy-MM <br/>
     * yyyy/MM <br/>
     * yyyy <br/>
     *
     * @param dataStr 日期字符串
     * @return 格式化之后的Timestamp对象
     */
    public static Timestamp praseDate(String dataStr) {
        try {
            Date date = null;

            //String dataStr = values.get(i).trim();
            String firstPart = null;

            if (dataStr.contains(" ")) {
                firstPart = dataStr.split(" ")[0];
            } else {
                firstPart = dataStr;
            }

            dataStr = dataStr.replaceAll("/", "-");

            String[] temps = firstPart.split("-");

            for (int j = 0; j < temps.length; j++) {
                if (j == 1 || j == 2) {
                    if (temps[j].length() < 2) {
                        temps[j] = "0" + temps[j];
                    }
                }
            }

            String firstPartAc = "";
            for (int j = 0; j < temps.length; j++) {
                firstPartAc += temps[j] + "-";
            }
            firstPartAc = firstPartAc.substring(0,
                    firstPartAc.length() - 1);

            dataStr = dataStr.replace(firstPart, firstPartAc);
            //noinspection Duplicates
            if (dataStr.contains(" ")) {
                date = converseSimpleDT("yyyy-MM-dd HH:mm:ss",
                        dataStr);
            } else {
                int part = dataStr.split("-").length;

                if (part == 1) {
                    date = converseSimpleDT("yyyy", dataStr);
                } else if (part == 2) {
                    date = converseSimpleDT("yyyy-MM", dataStr);
                } else if (part == 3) {
                    date = converseSimpleDT("yyyy-MM-dd", dataStr);
                }
            }
            assert date != null;
            return new Timestamp(date.getTime());
        } catch (Exception e) {
            e.printStackTrace();
            Timestamp time = new Timestamp(Date.parse(dataStr));
            //sqlCmd.setTime(tp, time);
            return time;
        }
    }

    /**
     * 转换格式
     *
     * @param pattern 日期格式
     * @param dateStr 日期字符串
     * @return Date对象
     * @throws java.text.ParseException 转换错误
     */
    public static Date converseSimpleDT(String pattern, String dateStr)
            throws java.text.ParseException {
        return new SimpleDateFormat(pattern, Locale.US).parse(dateStr);
    }


    /**
     * 根据单元格数据类型转换为对应类型的数据
     *
     * @param cell 单元格
     * @return 对应类型的单元格数据
     */
    public static Object transformCell(Cell cell) {
        Object cellValue = null;
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case ERROR:
                cellValue = cell.getRichStringCellValue().getString().trim();
                break;
            case BLANK:
                cellValue = cell.getRichStringCellValue().getString().trim();
                break;
            case STRING:
                cellValue = cell.getRichStringCellValue().getString().trim();
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    cellValue = cell.getDateCellValue();
                } else {
                    cellValue = new DecimalFormat("#").format(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                try {
                    cellValue = cell.getNumericCellValue();
                } catch (Exception e) {
                    cellValue = String.valueOf(cell.getRichStringCellValue());
                }
                break;
            default:
                cellValue = null;
        }
        return cellValue;
    }


    /**
     * 判断单元格类型转换为字符串
     *
     * @param cell 单元格
     * @return 单元格内容
     */
    private String getCellValue(Cell cell) {
        String cellValue = null;
        DecimalFormat df = new DecimalFormat("#");
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    cellValue = cell.getRichStringCellValue().getString().trim();
                    break;
                case NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        cellValue = sdf.format(cell.getDateCellValue());
                    } else {
                        cellValue = df.format(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
                    break;
                case FORMULA:
                    try {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e) {
                        cellValue = String.valueOf(cell.getRichStringCellValue());
                    }
                    break;
                default:
                    cellValue = null;
            }
        }
        if (null != cellValue && "".equals(cellValue)) {
            cellValue = null;
        }

        return cellValue;
    }


    /**
     * 获取单元格的内容,并转换成字符串
     *
     * @param cell 单元格
     * @return 单元格内容字符串
     */
    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return cell.getNumericCellValue() + "";
            case BLANK:
                return null;
            case BOOLEAN:
                return cell.getBooleanCellValue() + "";
            case ERROR:
                return cell.getErrorCellValue() + "";
            case FORMULA:
                return cell.getCellFormula();
            default:
                return null;
        }
    }

    /**
     * 获取单元格的类型
     *
     * @param cell 旧单元格
     * @return 单元格类型字符串
     */
    public static String getCellTypeAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return "CELL_TYPE_STRING" + "-文本";
            case NUMERIC:
                return "CELL_TYPE_NUMERIC" + "-数字";
            case BLANK:
                return "CELL_TYPE_BLANK" + "-空白";
            case BOOLEAN:
                return "CELL_TYPE_BOOLEAN" + "-布尔值";
            case ERROR:
                return "CELL_TYPE_ERROR" + "-错误";
            case FORMULA:
                return "CELL_TYPE_FORMULA" + "-公式";
            default:
                return "CELL_TYPE_UNKNOWN" + "-未知";
        }
    }

}
