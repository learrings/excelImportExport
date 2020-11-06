package com.example.excelImportExport.util.excel;

import com.google.common.collect.Lists;
import com.monitorjbl.xlsx.StreamingReader;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.util.CollectionUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.text.ParseException;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.UUID;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.LongSupplier;

@Slf4j
public class ExcelMain {

    /**
     * 创建导入工具类，为后续处理数据做初始化准备
     *
     * @param file                - 流文件
     * @param excelImportStrategy - 导入策略
     */
    public static <T> ExcelImport<T> createExcelImport(MultipartFile file, ExcelImportStrategy<T> excelImportStrategy) {
        return new ExcelImport<>(file, excelImportStrategy);
    }

    /**
     * 创建导出工具类，为后续处理数据做初始化准备
     *
     * @param excelExportStrategy - 导出策略
     */
    public static ExcelExport createExcelExport(ExcelExportStrategy excelExportStrategy) {
        return createExcelExport(excelExportStrategy, null, null);
    }

    /**
     * 创建导出工具类，为后续处理数据做初始化准备
     *
     * @param excelExportStrategy - 导出策略
     * @param response            -
     * @param fileName            - 导出文件名
     */
    public static ExcelExport createExcelExport(ExcelExportStrategy excelExportStrategy, HttpServletResponse response, String fileName) {

        if (response != null) {
            try {
                response.setCharacterEncoding(StandardCharsets.UTF_8.name());
                response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                response.setHeader("Content-Disposition", "attachment;filename=" +
                        new String(fileName.getBytes(StandardCharsets.UTF_8.name()), StandardCharsets.ISO_8859_1.name())
                        + ExcelCommon.EXCEL_SUFFIX);
            } catch (Exception e) {
                throw new RuntimeException(e.getMessage());
            }
        }

        return new ExcelExport(excelExportStrategy);
    }

    /**
     * 导入导出基础类
     */
    private static abstract class ExcelBase {

        /**
         * 是否初始化构建完毕
         */
        protected boolean buildFinish = false;

        protected Workbook workbook;

        public abstract ExcelBase build();

        /**
         * 是否初始化构建完毕
         */
        private void verify() {
            if (!this.buildFinish) {
                throw new RuntimeException("Please build(), First");
            }
        }

        /**
         * 通过workbook获得数据
         *
         * @param function - Workbook的函数式获取 如：(workbook)->workbook.getSheetAt(0).getSheetName()
         */
        public <T> T getByWorkbook(Function<Workbook, T> function) {
            this.verify();
            return function.apply(this.workbook);
        }

    }

    /**
     * 导入工具对象
     */
    public static class ExcelImport<T> extends ExcelBase {

        /**
         * 创建导入工具类
         *
         * @param file                - 流文件
         * @param excelImportStrategy - 导入策略
         */
        private ExcelImport(MultipartFile file, ExcelImportStrategy<T> excelImportStrategy) {

            try {
                this.originalFilename = file.getOriginalFilename();

                super.workbook = StreamingReader.builder()
                        .rowCacheSize(ExcelCommon.EXCEL_IMPORT_STREAMING_READER_ROW_CACHE_SIZE)
                        .bufferSize(ExcelCommon.EXCEL_IMPORT_STREAMING_READER_BUFFER_SIZE)
                        .open(file.getInputStream());

                this.excelImportStrategy = excelImportStrategy;
            } catch (Exception e) {
                throw new RuntimeException(e.getMessage());
            }
        }

        /**
         * Excel导入文件名称
         */
        private final String originalFilename;

        /**
         * 导入策略
         */
        private final ExcelImportStrategy<T> excelImportStrategy;

        @Override
        public ExcelImport<T> build() {

            if (this.excelImportStrategy == null) {
                throw new RuntimeException("excelImportStrategy must be not null");
            } else if (this.excelImportStrategy.getClazz() == null
                    || this.excelImportStrategy.getFiledArray() == null) {
                throw new RuntimeException("excelImportStrategy[class/filedArray] must be not null");
            }

            this.buildFinish = true;

            return this;

        }

        /**
         * 获取Excel导入文件名称
         */
        public String getOriginalFilename() {
            super.verify();
            return this.originalFilename;
        }

        /**
         * 解析List数据
         */
        public List<T> getList() {

            super.verify();

            List<T> result = Lists.newArrayList();

            try {

                // 循环Sheet获取
                for (int i = 0; i < super.workbook.getNumberOfSheets(); i++) {

                    Sheet sheet = super.workbook.getSheetAt(i);

                    if (!this.excelImportStrategy.getSpecifySheetNumList().contains(i)) {
                        continue;
                    }

                    log.info("Sheet info: name[{}],total[{}]", sheet.getSheetName(),
                            sheet.getLastRowNum() - this.excelImportStrategy.getBeginRowIndex() - 1);

                    this.analyseSheet(sheet, i, result);

                }

            } catch (Exception e) {
                log.warn("analyse row exception==>", e);
                throw new RuntimeException(e.getMessage());
            }

            return result;
        }

        /**
         * 解析sheet
         *
         * @param sheet           -
         * @param currentSheetNum - 当前sheet页的索引
         * @param result          - 返回结果
         */
        private void analyseSheet(Sheet sheet, int currentSheetNum, List<T> result)
                throws InvocationTargetException, IntrospectionException, InstantiationException, IllegalAccessException, ParseException {

            for (Row row : sheet) {

                // 从指定Excel[行]索引后开始解析
                if (row.getRowNum() < this.excelImportStrategy.getBeginRowIndex()) {
                    continue;
                }

                this.analyseRow(row, currentSheetNum, result);

            }
        }

        /**
         * 解析row
         *
         * @param row             -
         * @param currentSheetNum -当前sheet页的索引
         * @param result          - 返回结果
         */
        private void analyseRow(Row row, int currentSheetNum, List<T> result)
                throws IllegalAccessException, InstantiationException, InvocationTargetException, IntrospectionException, ParseException {

            T obj = this.excelImportStrategy.getClazz().newInstance();
            result.add(obj);

            ExcelImportStrategy.ExcelImportCellInfo cellInfo;

            for (Cell cell : row) {

                // 从指定Excel[列]索引后开始解析
                if (cell.getColumnIndex() < this.excelImportStrategy.getBeginColIndex()) {
                    continue;
                }

                cellInfo = new ExcelImportStrategy.ExcelImportCellInfo(
                        this.excelImportStrategy.getFiledArray()[cell.getColumnIndex() - this.excelImportStrategy.getBeginColIndex()],
                        currentSheetNum, cell
                );

                this.analyseCell(cellInfo, obj);
            }
        }

        /**
         * 解析Cell
         *
         * @param cellInfo-
         * @param obj-
         */
        private void analyseCell(ExcelImportStrategy.ExcelImportCellInfo cellInfo, T obj)
                throws IllegalAccessException, InstantiationException, InvocationTargetException, IntrospectionException, ParseException {

            PropertyDescriptor propertyDescriptor;
            Object invokeObject;

            // 复合属性&普通属性
            if (cellInfo.getFiled().contains(ExcelCommon.SEPARATOR_SPOT)) {

                String[] filedSplit = cellInfo.getFiled().split(ExcelCommon.SEPARATOR_SPOT);
                PropertyDescriptor filedPropertyDescriptor = new PropertyDescriptor(filedSplit[0], this.excelImportStrategy.getClazz());
                Class<?> beanClass = filedPropertyDescriptor.getPropertyType();
                Method readMethod = filedPropertyDescriptor.getReadMethod();
                invokeObject = readMethod.invoke(obj);
                if (invokeObject == null) {
                    Method filedWriteMethod = filedPropertyDescriptor.getWriteMethod();
                    filedWriteMethod.invoke(obj, beanClass.newInstance());
                    invokeObject = readMethod.invoke(obj);
                }
                propertyDescriptor = new PropertyDescriptor(filedSplit[1], beanClass);

            } else {
                invokeObject = obj;
                propertyDescriptor = new PropertyDescriptor(cellInfo.getFiled(), this.excelImportStrategy.getClazz());
            }

            Method writeMethod = propertyDescriptor.getWriteMethod();
            cellInfo.setPropertyType(propertyDescriptor.getPropertyType());

            Object cellValue = this.getCellValue(cellInfo);
            if (cellValue != null) {
                writeMethod.invoke(invokeObject, cellValue);
            }
        }

        /**
         * 获取单元格值
         *
         * @param cellInfo-
         */
        private Object getCellValue(ExcelImportStrategy.ExcelImportCellInfo cellInfo) throws ParseException {

            this.buildValue(cellInfo);

            if (Objects.isNull(cellInfo.getCellValue())) {
                return null;
            }

            // 执行策略规则
            List<Consumer<ExcelImportStrategy.ExcelImportCellInfo>> strategyList =
                    this.excelImportStrategy.getRuleMap().get(cellInfo.getFiled());

            if (!CollectionUtils.isEmpty(strategyList)) {
                for (Consumer<ExcelImportStrategy.ExcelImportCellInfo> strategy : strategyList) {
                    strategy.accept(cellInfo);
                }
            }

            // 转换为属性对应的类型
            if (cellInfo.getPropertyType() == Byte.class || cellInfo.getPropertyType() == byte.class) {
                return Byte.parseByte(cellInfo.getCellValue().toString());
            } else if (cellInfo.getPropertyType() == short.class || cellInfo.getPropertyType() == Short.class) {
                return Short.parseShort(cellInfo.getCellValue().toString());
            } else if (cellInfo.getPropertyType() == Integer.class || cellInfo.getPropertyType() == int.class) {
                return Integer.parseInt(cellInfo.getCellValue().toString());
            } else if (cellInfo.getPropertyType() == long.class || cellInfo.getPropertyType() == Long.class) {
                return Long.parseLong(cellInfo.getCellValue().toString());
            } else if (cellInfo.getPropertyType() == double.class || cellInfo.getPropertyType() == Double.class) {
                return Double.parseDouble(cellInfo.getCellValue().toString());
            } else if (cellInfo.getPropertyType() == BigDecimal.class) {
                return new BigDecimal(cellInfo.getCellValue().toString());
            } else if (cellInfo.getPropertyType() == float.class || cellInfo.getPropertyType() == Float.class) {
                return Float.parseFloat(cellInfo.getCellValue().toString());
            } else if (!(cellInfo.getCellValue() instanceof Date) && cellInfo.getPropertyType() == Date.class) {
                return ExcelCommon.getNormalDate(cellInfo.getCellValue().toString());
            }

            return cellInfo.getCellValue();
        }

        /**
         * 通过poi初始化构建值
         *
         * @param cellInfo-
         */
        private void buildValue(ExcelImportStrategy.ExcelImportCellInfo cellInfo) {

            Cell cell = cellInfo.getCell();

            switch (cell.getCellTypeEnum()) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellInfo.setCellValue(cell.getDateCellValue());
                    } else {
                        cellInfo.setCellValue(NumberToTextConverter.toText(cell.getNumericCellValue()));
                    }
                    break;
                case STRING:
                    cellInfo.setCellValue(cell.getRichStringCellValue().getString());
                    break;
                case FORMULA:
                    cellInfo.setCellValue(cell.getCellFormula());
                    break;
                case BLANK:
                    break;
                case BOOLEAN:
                    cellInfo.setCellValue(cell.getBooleanCellValue());
                case ERROR:
                    throw new RuntimeException("cell is error value");
                default:
                    throw new RuntimeException("cell is unknown type");
            }
        }

    }

    /**
     * 导出工具对象
     */
    public static class ExcelExport extends ExcelBase {

        /**
         * 创建导出工具类
         *
         * @param excelExportStrategy - 导出策略
         */
        private ExcelExport(ExcelExportStrategy excelExportStrategy) {
            this.excelExportStrategy = excelExportStrategy;
        }

        /**
         * 当前sheet页对象
         */
        private Sheet sheet;

        /**
         * 文件本地存储位置
         */
        private String fileTempPath;

        /**
         * 输出流
         */
        private OutputStream fileOutputStream;

        /**
         * 导出策略
         */
        private final ExcelExportStrategy excelExportStrategy;

        /**
         * sheet页名称前缀
         */
        private String sheetNamePrefix;

        /**
         * 单个sheet页存放最大数据量，超过则加sheet页
         */
        private Integer sheetMaxRow;

        /**
         * sheet页配置（自定义）
         *
         * @param sheetNamePrefix - sheet页名称前缀
         * @param sheetMaxRow     - 单个sheet页存放最大数据量，超过则加sheet页
         */
        public ExcelExport setSheetConfig(String sheetNamePrefix, Integer sheetMaxRow) {

            if (StringUtils.isNotBlank(sheetNamePrefix)) {
                this.sheetNamePrefix = sheetNamePrefix;
            }

            if (sheetMaxRow != null && sheetMaxRow > 0) {
                this.sheetMaxRow = sheetMaxRow;
            }

            return this;
        }

        /**
         * 输出流设置（自定义）
         *
         * @param servletOutputStream -
         */
        public ExcelExport setFileOutputStream(ServletOutputStream servletOutputStream) {
            this.fileOutputStream = servletOutputStream;
            return this;
        }

        @Override
        public ExcelExport build() {

            try {

                if (this.excelExportStrategy == null) {
                    throw new RuntimeException("excelExportStrategy must be not null");
                } else if (this.excelExportStrategy.getCellList() == null) {
                    throw new RuntimeException("excelExportStrategy[cellList] must be not null");
                }

                super.workbook = new SXSSFWorkbook();

                if (this.sheetNamePrefix == null) {
                    this.sheetNamePrefix = ExcelCommon.EXCEL_EXPORT_SHEET_NAME_DEFAULT;
                }

                if (this.sheetMaxRow == null) {
                    this.sheetMaxRow = ExcelCommon.EXCEL_EXPORT_SHEET_MAX_ROW;
                }

                this.buildSheet();

                if (this.fileOutputStream == null) {
                    this.fileTempPath = StringUtils.join(System.getProperty(TempFile.JAVA_IO_TMPDIR), UUID.randomUUID(), ExcelCommon.EXCEL_SUFFIX);
                    this.fileOutputStream = new FileOutputStream(this.fileTempPath.replace("//", "/"));
                }

                this.buildFinish = true;

            } catch (Exception e) {
                throw new RuntimeException(e.getMessage());
            }

            return this;
        }

        /**
         * 导出List数据
         */
        public String exportList() {

            super.verify();

            long beginId = 0L;

            for (; ; ) {

                // 只导出标题
                if (this.excelExportStrategy.getOption() == null) {
                    break;
                }

                List<? extends LongSupplier> list = this.excelExportStrategy.getOption().apply(beginId);

                if (!this.renderingRow(list)) {
                    break;
                }

                long maxId = list.stream().mapToLong(LongSupplier::getAsLong).max()
                        .orElseThrow(() -> new RuntimeException("not find max id"));

                // 无限循环，报错处理
                if (beginId >= maxId) {
                    throw new RuntimeException("have an infinite loop");
                }

                beginId = maxId;

            }

            try {
                super.workbook.write(this.fileOutputStream);
                if (this.fileOutputStream instanceof ServletOutputStream) {
                    this.fileOutputStream.flush();
                } else {
                    this.fileOutputStream.close();
                }
            } catch (IOException e) {
                log.warn("IOException", e);
                throw new RuntimeException(e);
            }

            return this.fileTempPath;
        }

        /**
         * 渲染行数据
         *
         * @param list -
         */
        private <T> boolean renderingRow(List<T> list) {

            if (CollectionUtils.isEmpty(list)) {
                return false;
            }

            for (T obj : list) {

                // 排除标题行之上的，只比较数据
                if ((this.sheet.getLastRowNum() - this.excelExportStrategy.getTitleBeginIndex() + 1) > this.sheetMaxRow) {
                    this.buildSheet();
                }

                Row row = this.sheet.createRow(this.sheet.getLastRowNum() + 1);
                this.renderingCell(row, obj);

            }

            return true;
        }

        /**
         * 渲染单元格数据
         *
         * @param row - 行对象
         * @param obj - 待渲染的数据对象
         */
        private <T> void renderingCell(Row row, T obj) {

            try {

                for (int i = 0; i < this.excelExportStrategy.getCellList().size(); i++) {

                    ExcelExportStrategy.ExcelCellModel cellModel = this.excelExportStrategy.getCellList().get(i);

                    Cell cell = row.createCell(i);

                    if (obj == null) {
                        // 渲染标题
                        cellModel.getTitleCell().accept(cell);
                        cell.setCellValue(cellModel.getTitleValue());
                    } else if (StringUtils.isNotBlank(cellModel.getFiled())) {
                        // 渲染属性值对应的单元格
                        Object value = new PropertyDescriptor(cellModel.getFiled(), obj.getClass()).getReadMethod().invoke(obj);
                        cell.setCellValue(cellModel.getDataValue().apply(cell, value));
                    }

                }
            } catch (Exception e) {
                throw new RuntimeException(e.getMessage());
            }

        }

        /**
         * 构建sheet页并初始化标题
         */
        private void buildSheet() {

            this.sheet = super.workbook.createSheet(this.sheetNamePrefix + super.workbook.getNumberOfSheets());
            Row row = this.sheet.createRow(this.excelExportStrategy.getTitleBeginIndex());
            this.renderingCell(row, null);

        }

    }


}
