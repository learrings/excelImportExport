package com.example.excelImportExport.dto.demo;

import lombok.Getter;
import lombok.Setter;

import java.util.Date;
import java.util.function.LongSupplier;

@Getter
@Setter
public class DemoDto implements LongSupplier {

    private Long id;
    private String name;
    private Date createTime;

    private Long beginId;

    @Override
    public long getAsLong() {
        // 导出必须实现（LongSupplier），设置主键，用于分页查询导出时，beginId使用
        return id;
    }
}
