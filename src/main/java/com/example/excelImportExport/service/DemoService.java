package com.example.excelImportExport.service;


import com.example.excelImportExport.dto.Pager;
import com.example.excelImportExport.dto.demo.DemoDto;

import java.util.List;

public interface DemoService {

    void importDemoDtoList(List<DemoDto> excelImportExportDtoList);

    List<DemoDto> queryPager(DemoDto excelImportExportDto, Pager pager);
}
