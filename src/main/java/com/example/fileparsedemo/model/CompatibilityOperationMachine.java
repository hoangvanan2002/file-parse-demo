package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CompatibilityOperationMachine {
    private ExcelCellValue machineCode;
    private ExcelCellValue priority;
    private ExcelCellValue altTransferMinute;
}
