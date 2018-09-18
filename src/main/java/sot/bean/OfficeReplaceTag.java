package sot.bean;
import org.apache.commons.lang3.StringUtils;

/**
 * 存放自定义标签数据的类
 * @author Monan
 * created on 9/17/2018 9:20 PM
 */
public final class OfficeReplaceTag {
    private OfficeReplaceTag() { }

    /**
     * 默认属性替换标签:{@code<SOF:xxxxx>}
     * 默认图片替换标签:{@code<SOF-IMG:xxxxx>}
     * @param propertyPreFix 属性标签前缀
     * @param propertySuffix 属性标签后缀
     * @param imgPreFix 图片标签前缀
     * @param imgSuffix 图片标签后缀
     */
    public OfficeReplaceTag(String propertyPreFix,String propertySuffix, String imgPreFix, String imgSuffix) {
        if (StringUtils.isNotBlank(propertyPreFix)) {
            this.propertyPreFix = propertyPreFix;
        } else {
            this.propertyPreFix = "<SOF:";
        }

        if (StringUtils.isNotBlank(propertySuffix)) {
            this.propertySuffix = propertySuffix;
        } else {
            this.propertySuffix = ">";
        }

        if (StringUtils.isNotBlank(imgPreFix)) {
            this.imgPreFix = imgPreFix;
        } else {
            this.imgPreFix = "<SOF-IMG:";
        }

        if (StringUtils.isNotBlank(imgSuffix)) {
            this.imgSuffix = imgSuffix;
        } else {
            this.imgSuffix = ">";
        }
    }

    /**
     * 属性标签前缀
     */
    private String propertyPreFix;
    /**
     * 属性标签后缀
     */
    private String propertySuffix;
    /**
     * 图片标签前缀
     */
    private String imgPreFix;
    /**
     * 图片标签后缀
     */
    private String imgSuffix;


    /**
     * Gets 图片标签后缀.
     *
     * @return Value of 图片标签后缀.
     */
    public String getImgSuffix() {
        return imgSuffix;
    }

    /**
     * Gets 图片标签前缀.
     *
     * @return Value of 图片标签前缀.
     */
    public String getImgPreFix() {
        return imgPreFix;
    }

    /**
     * Gets 属性标签后缀.
     *
     * @return Value of 属性标签后缀.
     */
    public String getPropertySuffix() {
        return propertySuffix;
    }

    /**
     * Gets 属性标签前缀.
     *
     * @return Value of 属性标签前缀.
     */
    public String getPropertyPreFix() {
        return propertyPreFix;
    }
}
