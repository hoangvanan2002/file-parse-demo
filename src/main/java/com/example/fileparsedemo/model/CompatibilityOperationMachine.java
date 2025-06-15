package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CompatibilityOperationMachine {
    private String machineCode;
    private String priority;
    private String altTransferMinute;
}
