package com.example.excelImportExport.util.excel;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.util.Assert;

import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;


/**
 * 导入策略对象
 */
@Getter
@Setter
public class ExcelImportStrategy<T> {

    /**
     * 导入策略定义
     * <p>
     * <br/>
     * 默认项：<br/>
     * this.beginRowIndex = 1;<br/>
     * this.beginColIndex = 0;<br/>
     * this.specifySheetNumList = Lists.newArrayList(0);<br/>
     *
     * @param clazz      - 转换对象
     * @param filedArray - 转换对象的属性组
     */
    public ExcelImportStrategy(Class<T> clazz, String[] filedArray) {
        this(clazz, filedArray, null, null);
    }

    /**
     * 导入策略定义，从第一列读取
     * <p>
     * <br/><br/>
     * 默认项：<br/>
     * this.beginRowIndex = 1;<br/>
     * this.beginColIndex = 0;<br/>
     * this.specifySheetNumList = Lists.newArrayList(0);<br/>
     *
     * @param clazz         - 转换对象
     * @param filedArray    - 转换对象的属性组
     * @param beginRowIndex - 从指定Excel[行]的索引位置读取，默认：1(排除标题行)
     * @param beginColIndex - 从指定Excel[列]的索引位置读取，默认：0
     */
    public ExcelImportStrategy(Class<T> clazz, String[] filedArray, Integer beginRowIndex, Integer beginColIndex) {
        Assert.notNull(clazz, "clazz is null");
        Assert.notNull(filedArray, "filedArray is null");

        this.clazz = clazz;

        this.filedArray = filedArray;

        if (beginRowIndex != null && beginRowIndex >= 0) {
            this.beginRowIndex = beginRowIndex;
        }

        if (beginColIndex != null && beginColIndex >= 0) {
            this.beginColIndex = beginColIndex;
        }

        this.beginRowIndex = 1;
        this.beginColIndex = 0;
        this.ruleMap = Maps.newHashMap();
        this.specifySheetNumList = Lists.newArrayList(0);
    }

    /**
     * 转换对象
     */
    private Class<T> clazz;

    /**
     * 转换对象的属性组
     */
    private String[] filedArray;

    /**
     * 起始读取Excel[行]的索引位置
     */
    private int beginRowIndex;

    /**
     * 起始读取Excel[列]的索引位置
     */
    private int beginColIndex;

    /**
     * 策略规则组<转换对象的属性,规则组>
     */
    private Map<String, List<Consumer<ExcelImportCellInfo>>> ruleMap;

    /**
     * 指定读取sheet页的索引位置，默认：只读取第一个sheet页
     */
    private List<Integer> specifySheetNumList;

    /**
     * 添加规则
     *
     * @param rule       - 导出之前按规则对值进行处理
     * @param filedArray - 可选，只针对指定的转换对象的属性有效，如果为空则对所有都有效
     */
    public ExcelImportStrategy<T> addRule(Consumer<ExcelImportCellInfo> rule, String... filedArray) {

        if (filedArray == null || filedArray.length == 0) {
            filedArray = this.filedArray;
        }

        for (String filed : filedArray) {
            List<Consumer<ExcelImportCellInfo>> strategyList = ruleMap.computeIfAbsent(filed, v -> Lists.newArrayList());
            strategyList.add(rule);
        }

        return this;
    }

    /**
     * 全局默认规则
     */
    public ExcelImportStrategy<T> addDefaultRule() {
        return this.addRule((cellInfo) -> {
        });
    }

    /**
     * 指定读取sheet页的索引位置（自定义）
     */
    public ExcelImportStrategy<T> addSpecifySheet(Integer... specifySheetNumArray) {

        if (specifySheetNumArray != null && specifySheetNumArray.length > 0) {
            specifySheetNumList.addAll(Arrays.asList(specifySheetNumArray));
        }

        return this;
    }


    /**
     * Excel读取的单元格信息，根据规则可以重新定义
     */
    @Getter
    @Setter
    public static class ExcelImportCellInfo {

        /**
         * 单元格信息
         *
         * @param filed           - 对象属性名
         * @param currentSheetNum - 当前sheet页索引
         * @param cell            - 当前单元格信息
         */
        public ExcelImportCellInfo(String filed, Integer currentSheetNum, Cell cell) {
            this.filed = filed;
            this.currentSheetNum = currentSheetNum;
            this.cell = cell;
        }

        /**
         * 转换对象的属性名
         */
        private String filed;

        /**
         * 当前单元格对应的sheet页索引
         */
        private Integer currentSheetNum;

        /**
         * 当前单元格信息
         */
        private Cell cell;

        /**
         * 转换对象的属性的类型
         */
        private Class<?> propertyType;

        /**
         * 当前单元格读取到的初始值
         */
        private Object cellValue;

    }
}
