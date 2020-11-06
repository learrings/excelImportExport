package com.example.excelImportExport.util.excel;

import com.google.common.collect.Lists;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;

import java.util.List;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.LongSupplier;

/**
 * 导出策略对象
 */
@Getter
@Setter
public class ExcelExportStrategy {

    /**
     * 导出策略定义
     * <p>
     * <br/>
     * 默认项：<br/>
     * this.titleBeginIndex = 0;<br/>
     *
     * @param option - 查询操作<查询结果最大Id(LongSupplier.getAsLong字段获取，初始化0), 查询结果List>
     */
    public ExcelExportStrategy(Function<Long, List<? extends LongSupplier>> option) {
        this.option = option;
        this.titleBeginIndex = 0;
    }

    /**
     * 导出循环查询操作，如
     * (beginId)->{return service.queryList(beginId)}
     */
    private Function<Long, List<? extends LongSupplier>> option;

    /**
     * 导出列配置信息
     */
    private List<ExcelCellModel> cellList;

    /**
     * 标题写入行索引，默认：0
     */
    private int titleBeginIndex;

    /**
     * 创建列基本信息
     *
     * @param filed - 属性名，可以为空
     * @param title - 标题名
     */
    public ExcelCellModel addCellModel(String filed, String title) {

        if (this.cellList == null) {
            this.cellList = Lists.newArrayList();
        }

        ExcelCellModel cellModel = new ExcelCellModel(filed, title);
        this.cellList.add(cellModel);
        return cellModel;
    }

    /**
     * 标题写入行索引（自定义）
     *
     * @param titleBeginIndex -
     */
    public void setTitleBeginIndex(int titleBeginIndex) {
        this.titleBeginIndex = titleBeginIndex;
    }

    @Getter
    public static class ExcelCellModel {

        private ExcelCellModel(String filed, String title) {
            this.filed = filed;
            this.titleValue = title;
            this.titleCell = (titleCell) -> {
            };
            this.dataValue = (cell, value) -> value.toString();
        }

        /**
         * 属性名
         */
        private final String filed;

        /**
         * 标题名
         */
        private final String titleValue;

        /**
         * 标题函数，用于自定义样式、备注等，默认：标准
         * 如(cell)->{}
         */
        private Consumer<Cell> titleCell;

        /**
         * 值函数，用于自定义值的变化方式/值样式等，默认：值的toString
         * 如(cell,dataValue)->{
         * // dataValue转换结果
         * return cellValueStr
         * }
         */
        private BiFunction<Cell, Object, String> dataValue;

        /**
         * 设置标题函数
         */
        public void setTitle(Consumer<Cell> titleCell) {
            this.titleCell = titleCell;
        }

        /**
         * 设置值函数
         */
        public void setData(BiFunction<Cell, Object, String> dataValue) {
            this.dataValue = dataValue;
        }
    }
}
