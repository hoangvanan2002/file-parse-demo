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
                    new String[]{"STT", "Bom Level", "Item Code", "Item Type", "Quantity", "Tỷ lệ hao hụt (%)", "Quy trình"},
                    result.getBomDetails(), (row, detail) -> {
                        row.createCell(0).setCellValue("");
                        if(detail.getBomLevel().equals("")) row.createCell(1).setCellValue("");
                        else row.createCell(1).setCellValue(Double.parseDouble(detail.getBomLevel()));
                        row.createCell(2).setCellValue(detail.getItemCode());
                        row.createCell(3).setCellValue(detail.getType());
                        if(detail.getQuantity().equals("")) row.createCell(1).setCellValue("");
                        else row.createCell(4).setCellValue(Double.parseDouble(detail.getQuantity()));
                        if(detail.getComponentYield().equals("")) row.createCell(1).setCellValue("");
                        else row.createCell(5).setCellValue(Double.parseDouble(detail.getComponentYield()));
                        row.createCell(6).setCellValue(detail.getTechnologyProcessCode());
                    });
            writeSheet(workbook, "operation", headerStyle,
                    new String[]{"Mã CĐ", "Tên CĐ", "Nhóm CĐ", "Nhân lực", "CT(s)", "Bộ phận", "Chức năng", "Mã tổ", "Tần suất chuyển đổi LOT", "Tỉ lệ hoàn thành CĐ", "Tỉ lệ vào/ra", "Leadtime"},
                    result.getOperations(),
                    (row, op) -> {
                        row.createCell(0).setCellValue(op.getOperationCode());
                        row.createCell(1).setCellValue(op.getOperationName());
                        row.createCell(2).setCellValue(op.getOperationGroup());
                        row.createCell(4).setCellValue(op.getCycleTime());
                        row.createCell(5).setCellValue(op.getDivisionId());
                        row.createCell(6).setCellValue(op.getEmployeeGroupCode());
                        row.createCell(7).setCellValue(op.getTransferFrequencyLot());
                        row.createCell(8).setCellValue(op.getCompletionRate());
                        row.createCell(9).setCellValue(op.getCompletionRate());
                        row.createCell(10).setCellValue(op.getInOutRatio());
                        row.createCell(11).setCellValue(op.getLeadTime());
                    });

            writeSheet(workbook, "technology_process", headerStyle,
                    new String[]{"Tên quy trình", "Mã quy trình"},
                    result.getTechnologyProcesses(),
                    (row, tp) -> {
                        row.createCell(0).setCellValue(tp.getTechnologyProcessName());
                        row.createCell(1).setCellValue(tp.getTechnologyProcessCode());
                    });

            writeSheet(workbook, "technology_process_operation", headerStyle,
                    new String[]{"Mã quy trình công nghệ", "Mã công đoạn", "Thứ tự công đoạn", "Mô tả", "Line"},
                    result.getTechnologyProcessOperations(),
                    (row, tpo) -> {
                        row.createCell(0).setCellValue(tpo.getTechnologyProcessCode());
                        row.createCell(1).setCellValue(tpo.getOperationCode());
                        row.createCell(2).setCellValue(tpo.getOperationOrder());
                        row.createCell(3).setCellValue(tpo.getDescription());
                        row.createCell(4).setCellValue(tpo.getOperationLine());
                    });

            writeSheet(workbook, "compatibility_operation_machine", headerStyle,
                    new String[]{"Mã máy", "Độ ưu tiên", "Thời gian di chuyển (phút)"},
                    result.getCompatibilityOperationMachines(),
                    (row, com) -> {
                        row.createCell(0).setCellValue(com.getMachineCode());
                        row.createCell(1).setCellValue(com.getPriority());
                        row.createCell(2).setCellValue(com.getAltTransferMinute());
                    });

            writeSheet(workbook, "products", headerStyle,
                    new String[]{"Mã hàng hóa", "Mã KH", "Tên tiếng Việt", "Tên tiếng Anh", "Dòng SP", "Loại", "Đặc tính GH", "Model", "Đơn vị"},
                    result.getProducts(),
                    (row, product) -> {
                        row.createCell(0).setCellValue(product.getProductCode());
                        row.createCell(1).setCellValue(product.getCustomerCode());
                        row.createCell(2).setCellValue(product.getProductName());
                        row.createCell(3).setCellValue(product.getProductEnName());
                        row.createCell(4).setCellValue(product.getProductLine());
                        row.createCell(5).setCellValue(product.getProductType());
                        row.createCell(6).setCellValue(product.getDeliveryCharacteristicCode());
                        row.createCell(7).setCellValue(product.getProductModel());
                        row.createCell(8).setCellValue(product.getUnit());
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
                String productCode = getCellValue(row, "B", evaluator);
                String bomLevel = getCellValue(row, "C", evaluator);
                String itemCode = getCellValue(row, "D", evaluator);
                String type = getCellValue(row, "E", evaluator);
                String quantity = getCellValue(row, "F", evaluator);
                String componentYield = getCellValue(row, "G", evaluator);
                String technologyProcessCode = getCellValue(row, "H", evaluator);

                if (hasAnyNonBlank(bomLevel, itemCode, type, quantity, componentYield, technologyProcessCode)) {
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
                String operationCode = getCellValue(row, "N", evaluator);
                String operationName = getCellValue(row, "O", evaluator);
                String operationGroup = getCellValue(row, "P", evaluator);
                String employeeQuantity = getCellValue(row, "Q", evaluator);
                String cycleTime = getCellValue(row, "R", evaluator);
                String divisionId = getCellValue(row, "S", evaluator);
                String employeeGroupCode = getCellValue(row, "U", evaluator);
                String transferFrequencyLot = getCellValue(row, "V", evaluator);
                String completionRate = getCellValue(row, "W", evaluator);
                String inOutRatio = getCellValue(row, "X", evaluator);
                String leadTime = getCellValue(row, "Y", evaluator);
                String machineGroupCode = getCellValue(row, "Z", evaluator);

                if (hasAnyNonBlank(operationCode, operationName, operationGroup, employeeQuantity, cycleTime, divisionId, employeeGroupCode, transferFrequencyLot, completionRate, inOutRatio, leadTime, machineGroupCode)) {
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
                String technologyProcessName = getCellValue(row, "I", evaluator);
                String technologyProcessCode1 = getCellValue(row, "J", evaluator);
                if (hasAnyNonBlank(technologyProcessName, technologyProcessCode1)) {
                    technologyProcessList.add(TechnologyProcess.builder()
                            .technologyProcessName(technologyProcessName)
                            .technologyProcessCode(technologyProcessCode1)
                            .build());
                }

                // TechnologyProcessOperation
                String technologyProcessCode2 = getCellValue(row, "J", evaluator);
                String operationCode2 = getCellValue(row, "K", evaluator);
                String operationOrder = getCellValue(row, "L", evaluator);
                String description = "";
                String operationLine = getCellValue(row, "H", evaluator);
                if (hasAnyNonBlank(technologyProcessCode2, operationCode2, operationOrder, operationLine)) {
                    technologyProcessOperationList.add(TechnologyProcessOperation.builder()
                            .technologyProcessCode(technologyProcessCode2)
                            .operationCode(operationCode2)
                            .operationOrder(operationOrder)
                            .description(description)
                            .operationLine(operationLine)
                            .build());
                }

                // CompatibilityOperationMachine
                String machineCode = getCellValue(row, "AA", evaluator);
                String priority = getCellValue(row, "AC", evaluator);
                String altTransferMinute = getCellValue(row, "AE", evaluator);
                if (hasAnyNonBlank(machineCode, priority, altTransferMinute)) {
                    compatibilityOperationMachineList.add(CompatibilityOperationMachine.builder()
                            .machineCode(machineCode)
                            .altTransferMinute(altTransferMinute)
                            .priority(priority)
                            .build());
                }

                // Products
                String productCode1 = getCellValue(row, "AG", evaluator);
                String customerCode = getCellValue(row, "AH", evaluator);
                String productName = getCellValue(row, "AI", evaluator);
                String productEnName = getCellValue(row, "AJ", evaluator);
                String productLine = getCellValue(row, "AK", evaluator);
                String productType = getCellValue(row, "AL", evaluator);
                String deliveryCharacteristicCode = getCellValue(row, "AM", evaluator);
                String productModel = getCellValue(row, "AN", evaluator);
                String productUnit = getCellValue(row, "AO", evaluator);

                if (hasAnyNonBlank(productCode1, customerCode, productName, productEnName, productLine, productType, deliveryCharacteristicCode, productModel, productUnit)) {
                    productsList.add(Products.builder()
                            .productCode(productCode1)
                            .customerCode(customerCode)
                            .productName(productName)
                            .productEnName(productEnName)
                            .productType(productType)
                            .deliveryCharacteristicCode(deliveryCharacteristicCode)
                            .productModel(productModel)
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
