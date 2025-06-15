package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TechnologyProcessOperation {
    private String technologyProcessCode;
    private String operationCode;
    private String operationOrder;
    private String description;
    private String operationLine;
}
