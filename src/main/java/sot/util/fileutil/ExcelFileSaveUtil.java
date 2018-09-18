package sot.util.fileutil;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;

/**
 * Excel文件读取工具类
 *
 * @author Monan
 * created on 2018/7/19 12:08
 */
public final class ExcelFileSaveUtil {
    /**
     * 私有化构造器
     */
    private ExcelFileSaveUtil() {
    }

    /**
     * 日志
     */
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelFileSaveUtil.class);

    /**
     * 使用指定名称保存Docx文件
     *
     * @param doc
     * @param path 保存路径
     * @param name 文件名
     * @return
     */
    @SuppressWarnings("ResultOfMethodCallIgnored")
    public static String saveXWPFDocumentUseSpecificName(XWPFDocument doc, String path, String name) {

        //检查数据有效性
        if (doc == null || StringUtils.isBlank(path)) {
            System.out.println("Error 保存路径为空");
            return null;
        }
        File checkFile = new File(path);
        if (!checkFile.exists()) {
            checkFile.mkdirs();
        }
        try {
            String filePath = path + name + ".docx";
            FileOutputStream fopts = new FileOutputStream(filePath);
            doc.write(fopts);
            fopts.flush();
            fopts.close();
            return filePath;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }


    /**
     * 保存Docx文件
     *
     * @param doc  Docx文件
     * @param path 保存路径
     * @return 返回文件绝对路径
     */
    @SuppressWarnings("ResultOfMethodCallIgnored")
    public static String saveXWPFDocument(XWPFDocument doc, String path) {

        //检查数据有效性
        if (doc == null || StringUtils.isBlank(path)) {
            System.out.println("Error 保存路径为空");
            return null;
        }
        File checkFile = new File(path);
        if (!checkFile.exists()) {
            checkFile.mkdirs();
        }
        try {
            String filePath = path + "POITest" + new Date().getTime() + ".docx";
            FileOutputStream fopts = new FileOutputStream(filePath);
            doc.write(fopts);
            fopts.flush();
            fopts.close();
            return filePath;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
