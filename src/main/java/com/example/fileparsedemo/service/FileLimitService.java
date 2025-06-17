package com.example.fileparsedemo.service;

import com.example.fileparsedemo.model.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;

@Service
public class FileLimitService {

    public byte[] writeExcelFile(MultipartFile file, ResultResponse result){
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.BLACK.getIndex());
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // Tạo style căn giữa
            CellStyle centerStyle = workbook.createCellStyle();
            centerStyle.setAlignment(HorizontalAlignment.CENTER);
            centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            writeSheet(workbook, "bom_detail", headerStyle,
                    new String[]{"STT", "Bom Level", "Item Code", "Item Type", "Quantity",
                            "Tỷ lệ hao hụt (%)", "Quy trình"},
                    result.getBomDetails(), (row, detail) -> {
                        row.createCell(0).setCellValue("");
                        row.createCell(1).setCellValue(detail.getBomLevel().getValue());
                        row.createCell(2).setCellValue(detail.getItemCode().getValue());
                        row.createCell(3).setCellValue(detail.getType().getValue());
                        row.createCell(4).setCellValue(detail.getQuantity().getValue());
                        row.createCell(5).setCellValue(detail.getComponentYield().getValue());
                        row.createCell(6).setCellValue(detail.getTechnologyProcessCode().getValue());
                    });
            writeSheet(workbook, "operation", headerStyle,
                    new String[]{"Mã CĐ", "Tên CĐ", "Nhóm CĐ", "Nhân lực", "CT(s)", "Bộ phận", "Chức năng", "Mã tổ", "Tần suất chuyển đổi LOT", "Tỉ lệ hoàn thành CĐ", "Tỉ lệ vào/ra", "Leadtime"},
                    result.getOperations(),
                    (row, op) -> {
                        row.createCell(0).setCellValue(op.getOperationCode().getValue());
                        row.createCell(1).setCellValue(op.getOperationName().getValue());
                        row.createCell(2).setCellValue(op.getOperationGroup().getValue());
                        row.createCell(3).setCellValue(op.getEmployeeQuantity().getValue());
                        row.createCell(4).setCellValue(op.getCycleTime().getValue());
                        row.createCell(5).setCellValue(op.getDivisionId().getValue());
                        row.createCell(6).setCellValue(op.getEmployeeGroupCode().getValue());
                        row.createCell(7).setCellValue(op.getTransferFrequencyLot().getValue());
                        row.createCell(8).setCellValue(op.getCompletionRate().getValue());
                        row.createCell(9).setCellValue(op.getCompletionRate().getValue());
                        row.createCell(10).setCellValue(op.getInOutRatio().getValue());
                        row.createCell(11).setCellValue(op.getLeadTime().getValue());
                    });

            writeSheet(workbook, "technology_process", headerStyle,
                    new String[]{"Tên quy trình", "Mã quy trình"},
                    result.getTechnologyProcesses(),
                    (row, tp) -> {
                        row.createCell(0).setCellValue(tp.getTechnologyProcessName().getValue());
                        row.createCell(1).setCellValue(tp.getTechnologyProcessCode().getValue());
                    });

            writeSheet(workbook, "technology_process_operation", headerStyle,
                    new String[]{"Mã quy trình công nghệ", "Mã công đoạn", "Thứ tự công đoạn", "Mô tả", "Line"},
                    result.getTechnologyProcessOperations(),
                    (row, tpo) -> {
                        row.createCell(0).setCellValue(tpo.getTechnologyProcessCode().getValue());
                        row.createCell(1).setCellValue(tpo.getOperationCode().getValue());
                        row.createCell(2).setCellValue(tpo.getOperationOrder().getValue());
                        row.createCell(3).setCellValue(tpo.getDescription().getValue());
                        row.createCell(4).setCellValue(tpo.getOperationLine().getValue());
                    });

            writeSheet(workbook, "compatibility_operation_machine", headerStyle,
                    new String[]{"Mã máy", "Độ ưu tiên", "Thời gian di chuyển (phút)"},
                    result.getCompatibilityOperationMachines(),
                    (row, com) -> {
                        row.createCell(0).setCellValue(com.getMachineCode().getValue());
                        row.createCell(1).setCellValue(com.getPriority().getValue());
                        row.createCell(2).setCellValue(com.getAltTransferMinute().getValue());
                    });

            writeSheet(workbook, "products", headerStyle,
                    new String[]{"Mã hàng hóa", "Mã KH", "Tên tiếng Việt", "Tên tiếng Anh", "Dòng SP", "Loại", "Đặc tính GH", "Model", "Đơn vị"},
                    result.getProducts(),
                    (row, product) -> {
                        row.createCell(0).setCellValue(product.getProductCode().getValue());
                        row.createCell(1).setCellValue(product.getCustomerCode().getValue());
                        row.createCell(2).setCellValue(product.getProductName().getValue());
                        row.createCell(3).setCellValue(product.getProductEnName().getValue());
                        row.createCell(4).setCellValue(product.getProductLine().getValue());
                        row.createCell(5).setCellValue(product.getProductType().getValue());
                        row.createCell(6).setCellValue(product.getDeliveryCharacteristicCode().getValue());
                        row.createCell(7).setCellValue(product.getProductModel().getValue());
                        row.createCell(8).setCellValue(product.getUnit().getValue());
                    });
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                workbook.write(out);
                return out.toByteArray(); // Trả kết quả dạng byte[]
            }

        } catch (Exception e) {
            e.printStackTrace();
            return new byte[0];
        }
    }

    private <T> void writeSheet(Workbook workbook, String sheetName, CellStyle headerStyle,
                                String[] headers, List<T> data, BiConsumer<Row, T> writer) {
        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
        // Data rows
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            writer.accept(row, data.get(i));
            for (int j = 0; j < headers.length; j++) {
                Cell cell = row.getCell(j);
                if (cell == null) cell = row.createCell(j);
                cell.setCellStyle(dataStyle);
            }
        }

        // Auto size tất cả cột
        int maxColumnIndex = 0;
        for (Row row : sheet) {
            if (row.getLastCellNum() > maxColumnIndex) {
                maxColumnIndex = row.getLastCellNum();
            }
        }
        for (int i = 0; i < maxColumnIndex; i++) {
            sheet.autoSizeColumn(i);
        }

    }

    public ResultResponse readExcelFile(MultipartFile file) throws Exception {
        List<BomDetail> bomDetails = new ArrayList<>();
        List<Operation> operations = new ArrayList<>();
        List<TechnologyProcess> technologyProcessList = new ArrayList<>();
        List<TechnologyProcessOperation> technologyProcessOperationList = new ArrayList<>();
        List<CompatibilityOperationMachine> compatibilityOperationMachineList = new ArrayList<>();
        List<Products> productsList = new ArrayList<>();
        try (InputStream is = file.getInputStream(); Workbook workbook = new XSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (int i = 3; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // BomDetail
                ExcelCellValue productCode = new ExcelCellValue("B", i, getCellValue(row, "B", evaluator));
                ExcelCellValue bomLevel= new ExcelCellValue("C", i, getCellValue(row, "C", evaluator));
                ExcelCellValue itemCode = new ExcelCellValue("D", i, getCellValue(row, "D", evaluator));
                ExcelCellValue type = new ExcelCellValue("E", i, getCellValue(row, "E", evaluator));
                ExcelCellValue quantity = new ExcelCellValue("F", i, getCellValue(row, "F", evaluator));
                ExcelCellValue componentYield = new ExcelCellValue("G", i, getCellValue(row, "G", evaluator));
                ExcelCellValue technologyProcessCode = new ExcelCellValue("H", i, getCellValue(row, "H", evaluator));
                if (hasAnyNonBlank(bomLevel.getValue(), itemCode.getValue(), type.getValue(), quantity.getValue(),
                        componentYield.getValue(), technologyProcessCode.getValue())) {
                    bomDetails.add(BomDetail.builder()
                            .productCode(productCode)
                            .bomLevel(bomLevel)
                            .itemCode(itemCode)
                            .type(type)
                            .quantity(quantity)
                            .componentYield(componentYield)
                            .technologyProcessCode(technologyProcessCode)
                            .build());
                }

                // Operation
                ExcelCellValue operationCode = new ExcelCellValue("N", i, getCellValue(row, "N", evaluator));
                ExcelCellValue operationName = new ExcelCellValue("O", i, getCellValue(row, "O", evaluator));
                ExcelCellValue operationGroup = new ExcelCellValue("P", i, getCellValue(row, "P", evaluator));
                ExcelCellValue employeeQuantity = new ExcelCellValue("Q", i, getCellValue(row, "Q", evaluator));
                ExcelCellValue cycleTime = new ExcelCellValue("R", i, getCellValue(row, "R", evaluator));
                ExcelCellValue divisionId = new ExcelCellValue("S", i, getCellValue(row, "S", evaluator));
                ExcelCellValue employeeGroupCode = new ExcelCellValue("U", i, getCellValue(row, "U", evaluator));
                ExcelCellValue transferFrequencyLot = new ExcelCellValue("V", i, getCellValue(row, "V", evaluator));
                ExcelCellValue completionRate = new ExcelCellValue("W", i, getCellValue(row, "W", evaluator));
                ExcelCellValue inOutRatio = new ExcelCellValue("X", i, getCellValue(row, "X", evaluator));
                ExcelCellValue leadTime = new ExcelCellValue("Y", i, getCellValue(row, "Y", evaluator));
                ExcelCellValue machineGroupCode = new ExcelCellValue("Z", i, getCellValue(row, "Z", evaluator));
                if (hasAnyNonBlank(operationCode.getValue(), operationName.getValue(), operationGroup.getValue(),
                        employeeQuantity.getValue(), cycleTime.getValue(), divisionId.getValue(), employeeGroupCode.getValue(),
                        transferFrequencyLot.getValue(), completionRate.getValue(), inOutRatio.getValue(),
                        leadTime.getValue(), machineGroupCode.getValue())) {
                    operations.add(Operation.builder()
                            .operationCode(operationCode)
                            .operationName(operationName)
                            .operationGroup(operationGroup)
                            .employeeQuantity(employeeQuantity)
                            .cycleTime(cycleTime)
                            .divisionId(divisionId)
                            .employeeGroupCode(employeeGroupCode)
                            .transferFrequencyLot(transferFrequencyLot)
                            .completionRate(completionRate)
                            .inOutRatio(inOutRatio)
                            .leadTime(leadTime)
                            .machineGroupCode(machineGroupCode)
                            .build());
                }

                // TechnologyProcess
                ExcelCellValue technologyProcessName = new ExcelCellValue("I", i, getCellValue(row, "I", evaluator));
                ExcelCellValue technologyProcessCode1 = new ExcelCellValue("J", i, getCellValue(row, "J", evaluator));
                if (hasAnyNonBlank(technologyProcessName.getValue(), technologyProcessCode1.getValue())) {
                    technologyProcessList.add(TechnologyProcess.builder()
                            .technologyProcessName(technologyProcessName)
                            .technologyProcessCode(technologyProcessCode1)
                            .build());
                }

                // TechnologyProcessOperation
                ExcelCellValue technologyProcessCode2 = new ExcelCellValue("J", i, getCellValue(row, "J", evaluator));
                ExcelCellValue operationCode2 = new ExcelCellValue("K", i, getCellValue(row, "K", evaluator));
                ExcelCellValue operationOrder = new ExcelCellValue("L", i, getCellValue(row, "L", evaluator));
                ExcelCellValue description = new ExcelCellValue("", i, "");
                ExcelCellValue operationLine = new ExcelCellValue("H", i, getCellValue(row, "H", evaluator));
                if (hasAnyNonBlank(technologyProcessCode2.getValue(), operationCode2.getValue(),
                        operationOrder.getValue(), operationLine.getValue())) {
                    technologyProcessOperationList.add(TechnologyProcessOperation.builder()
                            .technologyProcessCode(technologyProcessCode2)
                            .operationCode(operationCode2)
                            .operationOrder(operationOrder)
                            .description(description)
                            .operationLine(operationLine)
                            .build());
                }

                // CompatibilityOperationMachine
                ExcelCellValue machineCode = new ExcelCellValue("AA", i, getCellValue(row, "AA", evaluator));
                ExcelCellValue priority = new ExcelCellValue("AC", i, getCellValue(row, "AC", evaluator));
                ExcelCellValue altTransferMinute = new ExcelCellValue("AE", i, getCellValue(row, "AE", evaluator));
                if (hasAnyNonBlank(machineCode.getValue(), priority.getValue(), altTransferMinute.getValue())) {
                    compatibilityOperationMachineList.add(CompatibilityOperationMachine.builder()
                            .machineCode(machineCode)
                            .altTransferMinute(altTransferMinute)
                            .priority(priority)
                            .build());
                }

                // Products
                ExcelCellValue productCode1 = new ExcelCellValue("AG", i, getCellValue(row, "AG", evaluator));
                ExcelCellValue customerCode = new ExcelCellValue("AH", i, getCellValue(row, "AH", evaluator));
                ExcelCellValue productName = new ExcelCellValue("AI", i, getCellValue(row, "AI", evaluator));
                ExcelCellValue productEnName = new ExcelCellValue("AJ", i, getCellValue(row, "AJ", evaluator));
                ExcelCellValue productLine = new ExcelCellValue("AK", i, getCellValue(row, "AK", evaluator));
                ExcelCellValue productType = new ExcelCellValue("AL", i, getCellValue(row, "AL", evaluator));
                ExcelCellValue deliveryCharacteristicCode = new ExcelCellValue("AM", i, getCellValue(row, "AM", evaluator));
                ExcelCellValue productModel = new ExcelCellValue("AN", i, getCellValue(row, "AN", evaluator));
                ExcelCellValue productUnit = new ExcelCellValue("AO", i, getCellValue(row, "AO", evaluator));
                if (hasAnyNonBlank(productCode1.getValue(), customerCode.getValue(), productName.getValue(),
                        productEnName.getValue(), productLine.getValue(), productType.getValue(),
                        deliveryCharacteristicCode.getValue(), productModel.getValue(), productUnit.getValue())) {
                    productsList.add(Products.builder()
                            .productCode(productCode1)
                            .customerCode(customerCode)
                            .productName(productName)
                            .productEnName(productEnName)
                            .productType(productType)
                            .deliveryCharacteristicCode(deliveryCharacteristicCode)
                            .productModel(productModel)
                            .productLine(productLine)
                            .unit(productUnit)
                            .build());
                }
            }
        }
        return ResultResponse.builder()
                .bomDetails(bomDetails)
                .operations(operations)
                .technologyProcesses(technologyProcessList)
                .technologyProcessOperations(technologyProcessOperationList)
                .compatibilityOperationMachines(compatibilityOperationMachineList)
                .products(productsList)
                .build();
    }

    private static int columnLetterToIndex(String column) {
        int result = 0;
        for (int i = 0; i < column.length(); i++) {
            result *= 26;
            result += column.charAt(i) - 'A' + 1;
        }
        return result - 1;
    }

    private String getCellValue(Row row, String columnLetter, FormulaEvaluator evaluator) {
        int colIndex = columnLetterToIndex(columnLetter);
        Cell cell = row.getCell(colIndex);
        if (cell == null) return "";

        DataFormatter formatter = new DataFormatter();
        if (cell.getCellTypeEnum() == CellType.FORMULA) {
            CellValue cellValue = evaluator.evaluate(cell);
            if (cellValue == null) return "";
            return switch (cellValue.getCellTypeEnum()) {
                case STRING -> cellValue.getStringValue();
                case NUMERIC -> String.valueOf(cellValue.getNumberValue());
                case BOOLEAN -> String.valueOf(cellValue.getBooleanValue());
                default -> "";
            };
        }
        return formatter.formatCellValue(cell).trim();
    }

    private boolean hasAnyNonBlank(String... values) {
        for (String v : values) {
            if (v != null && !v.isBlank()) return true;
        }
        return false;
    }
}
