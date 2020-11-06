package com.example.excelImportExport.dto;

import com.example.excelImportExport.config.SpringContext;
import com.example.excelImportExport.dto.demo.DemoDto;
import com.example.excelImportExport.service.DemoService;
import com.example.excelImportExport.util.excel.ExcelExportStrategy;
import com.example.excelImportExport.util.excel.ExcelImportStrategy;
import org.apache.commons.lang3.time.DateFormatUtils;

import java.util.Date;
import java.util.List;
import java.util.function.Function;
import java.util.function.LongSupplier;

/**
 * 模拟创建策略
 */
public class ImportStrategyManage {

    /**
     * 模拟创建导入策略 DemoDto
     */
    public static ExcelImportStrategy<DemoDto> createImportDemoDto() {

        // 定义策略对象
        ExcelImportStrategy<DemoDto> excelImportStrategy = new ExcelImportStrategy<>(DemoDto.class,
                new String[]{"id", "name", "createTime"});

        // 添加策略规则（非必须），不加规则相当于属性值字符串展示
        excelImportStrategy
                // 全部参数适用
                .addRule((cellInfo) -> {

                            if ("11".equals(cellInfo.getCellValue())) {
                                cellInfo.setCellValue("22");
                            }
                        }
                )
                // id,name参数适用
                .addRule((cellInfo) -> {
                            if ("22".equals(cellInfo.getCellValue())) {
                                cellInfo.setCellValue("44");
                            }
                        }, "id", "name"
                )
                // name参数适用
                .addRule((cellInfo) -> {
                            if ("33".equals(cellInfo.getCellValue())) {
                                cellInfo.setCellValue("55");
                            }
                        }, "name"
                )
                // 默认全局规则（非必须）
                .addDefaultRule();

        excelImportStrategy.setBeginColIndex(1);

        return excelImportStrategy;
    }

    /**
     * 模拟创建导出策略 DemoDto
     *
     * @param demoDto -
     */
    public static ExcelExportStrategy createExportDemoDto(DemoDto demoDto) {

        Function<Long, List<? extends LongSupplier>> option = null;

        // 如果为空，相当于导出模板 无数据
        if (demoDto != null) {
            option = (beginId) -> {
                demoDto.setBeginId(beginId + 1);
                return SpringContext.getBean(DemoService.class).queryPager(demoDto, new Pager());
            };
        }


        ExcelExportStrategy excelExportStrategy = new ExcelExportStrategy(option);

        excelExportStrategy.addCellModel(null, "原因");
        excelExportStrategy.addCellModel("id", "主键");
        excelExportStrategy.addCellModel("name", "名称");
        excelExportStrategy.addCellModel("createTime", "创建日期")
                // 设置数据展示方式
                .setData((cell, value) ->
                        DateFormatUtils.format((Date) value, DateFormatUtils.ISO_8601_EXTENDED_DATE_FORMAT.getPattern())
                );

        return excelExportStrategy;
    }
}
