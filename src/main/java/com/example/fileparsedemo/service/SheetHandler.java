package com.example.fileparsedemo.service;

import com.example.fileparsedemo.model.*;
import lombok.RequiredArgsConstructor;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.util.*;

@RequiredArgsConstructor
public class SheetHandler extends DefaultHandler {

    private final SharedStringsTable sharedStringsTable;

    private final List<BomDetail> bomDetails = new ArrayList<>();
    private final List<Operation> operations = new ArrayList<>();
    private final List<TechnologyProcess> technologyProcesses = new ArrayList<>();
    private final List<TechnologyProcessOperation> technologyProcessOperations = new ArrayList<>();
    private final List<CompatibilityOperationMachine> compatibilityOperationMachines = new ArrayList<>();
    private final List<Products> products = new ArrayList<>();

    private final StringBuilder lastContents = new StringBuilder();
    private boolean nextIsString;
    private String currentColumn = "";
    private int currentRowIndex = -1;
    private final Map<String, ExcelCellValue> currentRowData = new HashMap<>();

    public List<BomDetail> getBomDetails() { return bomDetails; }
    public List<Operation> getOperations() { return operations; }
    public List<TechnologyProcess> getTechnologyProcesses() { return technologyProcesses; }
    public List<TechnologyProcessOperation> getTechnologyProcessOperations() { return technologyProcessOperations; }
    public List<CompatibilityOperationMachine> getCompatibilityOperationMachines() { return compatibilityOperationMachines; }
    public List<Products> getProducts() { return products; }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        if ("row".equals(name)) {
            currentRowData.clear();
            currentRowIndex = parseRowIndex(attributes.getValue("r"));
        }

        if ("c".equals(name)) {
            nextIsString = "s".equals(attributes.getValue("t"));
            currentColumn = extractColumnLetter(attributes.getValue("r"));
        }

        lastContents.setLength(0); // reset nội dung
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        lastContents.append(ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String name) {
        if ("v".equals(name)) {
            String value = nextIsString
                    ? new XSSFRichTextString(sharedStringsTable.getEntryAt(Integer.parseInt(lastContents.toString()))).toString()
                    : lastContents.toString();
            currentRowData.put(currentColumn, new ExcelCellValue(currentRowIndex, currentColumn, value.trim()));
        }

        if ("row".equals(name) && currentRowIndex >= 3) {
            handleBomDetailRow(currentRowData);
            handleOperationRow(currentRowData);
            handleTechnologyProcessRow(currentRowData);
            handleTechnologyProcessOperationRow(currentRowData);
            handleCompatibilityOperationMachineRow(currentRowData);
            handleProductsRow(currentRowData);
        }
    }

    // ----------------- Helper Methods ------------------
    private int parseRowIndex(String rowRef) {
        try {
            return rowRef == null ? -1 : Integer.parseInt(rowRef.replaceAll("[^0-9]", "")) - 1;
        } catch (NumberFormatException e) {
            return -1;
        }
    }

    private String extractColumnLetter(String cellRef) {
        return cellRef == null ? "" : cellRef.replaceAll("[^A-Z]", "");
    }

    private void handleBomDetailRow(Map<String, ExcelCellValue> row) {
        ExcelCellValue productCode = row.getOrDefault("B", null);
        ExcelCellValue bomLevel = row.getOrDefault("C", null);
        ExcelCellValue itemCode = row.getOrDefault("D", null);
        ExcelCellValue type = row.getOrDefault("E", null);
        ExcelCellValue quantity = row.getOrDefault("F", null);
        ExcelCellValue componentYield = row.getOrDefault("G", null);
        ExcelCellValue technologyProcessCode = row.getOrDefault("H", null);

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
    }

    private void handleOperationRow(Map<String, ExcelCellValue> row) {
        ExcelCellValue operationCode = row.getOrDefault("N", null);
        ExcelCellValue operationName = row.getOrDefault("O", null);
        ExcelCellValue operationGroup = row.getOrDefault("P", null);
        ExcelCellValue employeeQuantity = row.getOrDefault("Q", null);
        ExcelCellValue cycleTime = row.getOrDefault("R", null);
        ExcelCellValue divisionId = row.getOrDefault("S", null);
        ExcelCellValue employeeGroupCode = row.getOrDefault("U", null);
        ExcelCellValue transferFrequencyLot = row.getOrDefault("V", null);
        ExcelCellValue completionRate = row.getOrDefault("W", null);
        ExcelCellValue inOutRatio = row.getOrDefault("X", null);
        ExcelCellValue leadTime = row.getOrDefault("Y", null);
        ExcelCellValue machineGroupCode = row.getOrDefault("Z", null);

        if (hasAnyNonBlank(operationCode.getValue(), operationName.getValue(), operationGroup.getValue(),
                employeeQuantity.getValue(), cycleTime.getValue(), divisionId.getValue(),
                employeeGroupCode.getValue(), transferFrequencyLot.getValue(), completionRate.getValue(),
                inOutRatio.getValue(), leadTime.getValue(), machineGroupCode.getValue())) {

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
    }

    private void handleTechnologyProcessRow(Map<String, ExcelCellValue> row) {
        ExcelCellValue technologyProcessName = row.getOrDefault("I", null);
        ExcelCellValue technologyProcessCode = row.getOrDefault("J", null);
        if(hasAnyNonBlank(technologyProcessName.getValue(), technologyProcessCode.getValue())){
            technologyProcesses.add(TechnologyProcess.builder()
                    .technologyProcessName(technologyProcessName)
                    .technologyProcessCode(technologyProcessCode)
                    .build());
        }
    }

    private void handleTechnologyProcessOperationRow(Map<String, ExcelCellValue> row) {
        ExcelCellValue technologyProcessCode = row.getOrDefault("J", null);
        ExcelCellValue operationCode = row.getOrDefault("K", null);
        ExcelCellValue operationOrder = row.getOrDefault("L", null);
        ExcelCellValue description = new ExcelCellValue("", null, "");
        ExcelCellValue operationLine = row.getOrDefault("H", null);
        if(hasAnyNonBlank(technologyProcessCode.getValue(), operationCode.getValue(),
                operationOrder.getValue(), description.getValue(), operationLine.getValue())){
            technologyProcessOperations.add(TechnologyProcessOperation.builder()
                    .technologyProcessCode(technologyProcessCode)
                    .operationCode(operationCode)
                    .operationOrder(operationOrder)
                    .description(description)
                    .operationLine(operationLine)
                    .build());
        }
    }

    private void handleCompatibilityOperationMachineRow(Map<String, ExcelCellValue> row) {
        ExcelCellValue machineCode = row.getOrDefault("AA", null);
        ExcelCellValue priority= row.getOrDefault("AC", null);
        ExcelCellValue altTransferMinute = row.getOrDefault("AE", null);
        if(hasAnyNonBlank(machineCode.getValue(), priority.getValue(), altTransferMinute.getValue())){
            compatibilityOperationMachines.add(CompatibilityOperationMachine.builder()
                    .machineCode(machineCode)
                    .altTransferMinute(altTransferMinute)
                    .priority(priority)
                    .build());
        }
    }

    private void handleProductsRow(Map<String, ExcelCellValue> row) {
        ExcelCellValue productCode = row.getOrDefault("AG", null);
        ExcelCellValue customerCode = row.getOrDefault("AH", null);
        ExcelCellValue productName = row.getOrDefault("AI", null);
        ExcelCellValue productEnName = row.getOrDefault("AJ", null);
        ExcelCellValue productLine = row.getOrDefault("AK", null);
        ExcelCellValue productType = row.getOrDefault("AL", null);
        ExcelCellValue deliveryCharacteristicCode = row.getOrDefault("AM", null);
        ExcelCellValue productModel = row.getOrDefault("AN", null);
        ExcelCellValue productUnit = row.getOrDefault("AO", null);
        if(hasAnyNonBlank(productCode.getValue(), customerCode.getValue(),
                productName.getValue(), productEnName.getValue(),
                productLine.getValue(), productType.getValue(),
                deliveryCharacteristicCode.getValue(), productModel.getValue())){
            products.add(Products.builder()
                    .productCode(productCode)
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

    private boolean hasAnyNonBlank(String... values) {
        for (String v : values) {
            if (v != null && !v.isBlank()) return true;
        }
        return false;
    }
}
