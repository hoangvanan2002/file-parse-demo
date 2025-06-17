package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Operation {
    private ExcelCellValue operationCode;
    private ExcelCellValue operationName;
    private ExcelCellValue operationGroup;
    private ExcelCellValue employeeQuantity;
    private ExcelCellValue cycleTime;
    private ExcelCellValue divisionId;
    private ExcelCellValue employeeGroupCode;
    private ExcelCellValue transferFrequencyLot;
    private ExcelCellValue completionRate;
    private ExcelCellValue inOutRatio;
    private ExcelCellValue leadTime;
    private ExcelCellValue machineGroupCode;
}
