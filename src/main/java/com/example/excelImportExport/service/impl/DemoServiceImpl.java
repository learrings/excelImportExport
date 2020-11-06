package com.example.excelImportExport.service.impl;

import com.example.excelImportExport.dto.Pager;
import com.example.excelImportExport.dto.demo.DemoDto;
import com.example.excelImportExport.service.DemoService;
import com.google.common.collect.Lists;
import org.apache.commons.lang3.time.DateUtils;
import org.springframework.stereotype.Service;

import java.util.Date;
import java.util.List;

@Service
public class DemoServiceImpl implements DemoService {


    @Override
    public void importDemoDtoList(List<DemoDto> excelImportExportDtoList) {

    }


    @Override
    public List<DemoDto> queryPager(DemoDto excelImportExportDto, Pager pager) {

        // 模拟数据

        if (excelImportExportDto.getBeginId() > 22) {
            return null;
        }

        List<DemoDto> list = Lists.newArrayList();
        for (long i = excelImportExportDto.getBeginId(); i < excelImportExportDto.getBeginId() + 10; i++) {
            DemoDto obj = new DemoDto();
            obj.setId(i);
            obj.setName("小米" + i);
            obj.setCreateTime(DateUtils.addDays(new Date(), Integer.parseInt(i + "")));
            list.add(obj);
        }

        return list;
    }

}
