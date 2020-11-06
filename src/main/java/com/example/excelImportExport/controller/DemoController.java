package com.example.excelImportExport.controller;

import com.example.excelImportExport.component.FastFileStorageClient;
import com.example.excelImportExport.component.StorePath;
import com.example.excelImportExport.dto.ImportStrategyManage;
import com.example.excelImportExport.dto.demo.DemoDto;
import com.example.excelImportExport.service.DemoService;
import com.example.excelImportExport.util.excel.ExcelExportStrategy;
import com.example.excelImportExport.util.excel.ExcelImportStrategy;
import com.example.excelImportExport.util.excel.ExcelMain;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import springfox.documentation.annotations.ApiIgnore;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

@Slf4j
@RestController
@Api(value = "excelImportExport")
@RequestMapping(value = "/excelImportExport")
public class DemoController {

    @Resource
    private DemoService excelImportExportService;
    @Resource
    private FastFileStorageClient fastFileStorageClient;


    @ApiOperation(value = "导入Excel")
    @PostMapping(value = "/{menu}/importExcel", consumes = "multipart/*", headers = "content-type=multipart/form-data")
    public Object importExcel(@ApiParam(value = "文件上传", required = true) MultipartFile file,
                              @ApiParam() @PathVariable String menu) {

        // 设置导入策略（自定义）
        ExcelImportStrategy<DemoDto> excelImportStrategy = ImportStrategyManage.createImportDemoDto();

        // 返回解析结果
        List<DemoDto> excelImportExportDtoList = ExcelMain.createExcelImport(file, excelImportStrategy).build().getList();

        excelImportExportService.importDemoDtoList(excelImportExportDtoList);
        return excelImportExportDtoList;
    }

    @ApiOperation(value = "异步导出Excel")
    @PostMapping(value = "/{menu}/exportExcel")
    public String exportExcel(@ApiIgnore() DemoDto excelImportExportDto, @PathVariable String menu) throws IOException {

        // 设置导出策略（自定义）
        ExcelExportStrategy excelExportStrategy = ImportStrategyManage.createExportDemoDto(excelImportExportDto);

        // 返回本地缓存地址
        String fileTempName = ExcelMain.createExcelExport(excelExportStrategy).build().exportList();

        // 上传
        File file = new File(fileTempName);
        StorePath storePath = fastFileStorageClient.uploadFile(new FileInputStream(file),
                file.length(), file.getName().substring(file.getName().lastIndexOf(".") + 1));

        return storePath.getPath();
    }

    @ApiOperation(value = "导出excel")
    @GetMapping(value = "/getExcelModel", headers = "Accept=application/octet-stream")
    public void getExcelModel(HttpServletResponse response, @ApiParam() @RequestParam() String name) throws IOException {

        // 设置导出策略（自定义）
        ExcelExportStrategy excelExportStrategy = ImportStrategyManage.createExportDemoDto(new DemoDto());

        // 返回本地缓存地址
        ExcelMain.createExcelExport(excelExportStrategy, response, name)
                .setFileOutputStream(response.getOutputStream()).build().exportList();
    }

}
