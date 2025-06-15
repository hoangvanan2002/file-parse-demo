package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Operation {
    private String operationCode;
    private String operationName;
    private String operationGroup;
    private String employeeQuantity;
    private String cycleTime;
    private String divisionId;
    private String employeeGroupCode;
    private String transferFrequencyLot;
    private String completionRate;
    private String inOutRatio;
    private String leadTime;
    private String machineGroupCode;
}
