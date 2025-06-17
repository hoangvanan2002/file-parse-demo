package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BomDetail {
    private ExcelCellValue productCode;
    private ExcelCellValue bomLevel;
    private ExcelCellValue itemCode;
    private ExcelCellValue type;
    private ExcelCellValue quantity;
    private ExcelCellValue componentYield;
    private ExcelCellValue technologyProcessCode;
}
