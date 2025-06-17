package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@Builder
public class ExcelCellValue {

    public ExcelCellValue(String column, Integer row, String value) {
        this.column = column;
        this.row = row;
        this.value = value;
    }

    private String column;
    private Integer row;
    private String value;

    public ExcelCellValue(int currentRowIndex, String currentColumn, String value) {
        this.column = currentColumn;
        this.row = currentRowIndex;
        this.value = value;
    }
}
