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
    private final Map<String, String> currentRowData = new HashMap<>();

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

            currentRowData.put(currentColumn, value.trim());
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

    private void handleBomDetailRow(Map<String, String> row) {
        String productCode = row.getOrDefault("B", "");
        String bomLevel = row.getOrDefault("C", "");
        String itemCode = row.getOrDefault("D", "");
        String type = row.getOrDefault("E", "");
        String quantity = row.getOrDefault("F", "");
        String componentYield = row.getOrDefault("G", "");
        String technologyProcessCode = row.getOrDefault("H", "");

        if (hasAnyNonBlank(bomLevel, itemCode, type, quantity,
                componentYield, technologyProcessCode)) {
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

    private void handleOperationRow(Map<String, String> row) {
        String operationCode = row.getOrDefault("N", "");
        String operationName = row.getOrDefault("O", "");
        String operationGroup = row.getOrDefault("P", "");
        String employeeQuantity = row.getOrDefault("Q", "");
        String cycleTime = row.getOrDefault("R", "");
        String divisionId = row.getOrDefault("S", "");
        String employeeGroupCode = row.getOrDefault("U", "");
        String transferFrequencyLot = row.getOrDefault("V", "");
        String completionRate = row.getOrDefault("W", "");
        String inOutRatio = row.getOrDefault("X", "");
        String leadTime = row.getOrDefault("Y", "");
        String machineGroupCode = row.getOrDefault("Z", "");

        if (hasAnyNonBlank(operationCode, operationName, operationGroup, employeeQuantity, cycleTime,
                divisionId, employeeGroupCode, transferFrequencyLot, completionRate,
                inOutRatio, leadTime, machineGroupCode)) {

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

    private void handleTechnologyProcessRow(Map<String, String> row) {
        String technologyProcessName = row.getOrDefault("I", "");
        String technologyProcessCode = row.getOrDefault("J", "");
        if(hasAnyNonBlank(technologyProcessName, technologyProcessCode)){
            technologyProcesses.add(TechnologyProcess.builder()
                    .technologyProcessName(technologyProcessName)
                    .technologyProcessCode(technologyProcessCode)
                    .build());
        }
    }

    private void handleTechnologyProcessOperationRow(Map<String, String> row) {
        String technologyProcessCode = row.getOrDefault("J", "");
        String operationCode = row.getOrDefault("K", "");;
        String operationOrder = row.getOrDefault("L", "");;
        String description = "";
        String operationLine = row.getOrDefault("H", "");;
        if(hasAnyNonBlank(technologyProcessCode, operationCode, operationOrder, description, operationLine)){
            technologyProcessOperations.add(TechnologyProcessOperation.builder()
                    .technologyProcessCode(technologyProcessCode)
                    .operationCode(operationCode)
                    .operationOrder(operationOrder)
                    .description(description)
                    .operationLine(operationLine)
                    .build());
        }
    }

    private void handleCompatibilityOperationMachineRow(Map<String, String> row) {
        String machineCode = row.getOrDefault("AA", "");
        String priority= row.getOrDefault("AC", "");
        String altTransferMinute = row.getOrDefault("AE", "");
        if(hasAnyNonBlank(machineCode, priority, altTransferMinute)){
            compatibilityOperationMachines.add(CompatibilityOperationMachine.builder()
                    .machineCode(machineCode)
                    .altTransferMinute(altTransferMinute)
                    .priority(priority)
                    .build());
        }
    }

    private void handleProductsRow(Map<String, String> row) {
        String productCode = row.getOrDefault("AG", "");
        String customerCode = row.getOrDefault("AH", "");
        String productName = row.getOrDefault("AI", "");
        String productEnName = row.getOrDefault("AJ", "");
        String productLine = row.getOrDefault("AK", "");
        String productType = row.getOrDefault("AL", "");
        String deliveryCharacteristicCode = row.getOrDefault("AM", "");
        String productModel = row.getOrDefault("AN", "");
        String productUnit = row.getOrDefault("AO", "");
        if(hasAnyNonBlank(productCode, customerCode, productName, productEnName,
                productLine, productType, deliveryCharacteristicCode, productModel)){
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
