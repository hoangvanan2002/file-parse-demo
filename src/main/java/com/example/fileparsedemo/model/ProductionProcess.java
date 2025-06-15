package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ProductionProcess {
    private String technologyProcessName;
    private String technologyProcessCode;
    private String operationCode;
    private String operationOrder;
    private String operationLine;
}
