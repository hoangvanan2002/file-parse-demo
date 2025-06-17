package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TechnologyProcess {
    private ExcelCellValue technologyProcessName;
    private ExcelCellValue technologyProcessCode;
}
