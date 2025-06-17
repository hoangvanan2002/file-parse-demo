package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TechnologyProcessOperation {
    private ExcelCellValue technologyProcessCode;
    private ExcelCellValue operationCode;
    private ExcelCellValue operationOrder;
    private ExcelCellValue description;
    private ExcelCellValue operationLine;
}
