package com.example.excelImportExport.dto;

import lombok.Getter;
import lombok.Setter;

/**
 * 模拟pager
 */
@Getter
@Setter
public class Pager {

    private int pageIndex = 1;

    private int pageSize = 2000;
}
