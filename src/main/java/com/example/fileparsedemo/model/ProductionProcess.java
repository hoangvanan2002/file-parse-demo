package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ProductionProcess {
    private ExcelCellValue technologyProcessName;
    private ExcelCellValue technologyProcessCode;
    private ExcelCellValue operationCode;
    private ExcelCellValue operationOrder;
    private ExcelCellValue operationLine;
}
