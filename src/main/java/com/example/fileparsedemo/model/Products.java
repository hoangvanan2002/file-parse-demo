package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Products {
    private ExcelCellValue productCode;
    private ExcelCellValue customerCode;
    private ExcelCellValue productName;
    private ExcelCellValue productEnName;
    private ExcelCellValue productLine;
    private ExcelCellValue productType;
    private ExcelCellValue deliveryCharacteristicCode;
    private ExcelCellValue productModel;
    private ExcelCellValue unit;
}
