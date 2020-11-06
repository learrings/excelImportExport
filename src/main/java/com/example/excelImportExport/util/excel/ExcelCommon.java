package com.example.excelImportExport.util.excel;

import org.apache.commons.lang3.StringUtils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 *
 */
public class ExcelCommon {

    /**
     * excel尾缀
     */
    public static String EXCEL_SUFFIX = ".xlsx";

    /**
     * 分隔符
     */
    public static String SEPARATOR_SPOT = ".";

    /**
     * 导入本地缓存行
     */
    public static int EXCEL_IMPORT_STREAMING_READER_ROW_CACHE_SIZE = 100;

    /**
     * 导入本地缓存大小
     */
    public static int EXCEL_IMPORT_STREAMING_READER_BUFFER_SIZE = 4096;

    /**
     * 导出默认sheet页名字
     */
    public static String EXCEL_EXPORT_SHEET_NAME_DEFAULT = "数据导出";

    /**
     * 每个sheet页默认导出数据条数
     */
    public static int EXCEL_EXPORT_SHEET_MAX_ROW = 4000000;

    /**
     * 支持大部分日期转换  如 05-13-2019、05/13/2019、2019-01-01 13:26:10 、2019-1-1 1:6:0 等
     */
    public static Date getNormalDate(String dateStr) throws ParseException {
        if (StringUtils.isBlank(dateStr)) {
            return null;
        }

        String[] dateGroup = StringUtils.splitByCharacterType(dateStr);
        StringBuilder sb = new StringBuilder();
        for (String s : dateGroup) {
            if (!StringUtils.isNumeric(s)) {
                continue;
            }
            if (s.length() == 4) {
                sb.insert(0, s);
            } else {
                if (s.length() == 1) {
                    sb.append("0");
                }
                sb.append(s);
            }
        }

        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        return sdf.parse(StringUtils.rightPad(sb.toString(), 14, "0"));
    }

}
