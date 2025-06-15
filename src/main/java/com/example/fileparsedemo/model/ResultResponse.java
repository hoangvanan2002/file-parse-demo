package com.example.fileparsedemo.model;

import lombok.*;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ResultResponse {
    private List<BomDetail> bomDetails;
    private List<Operation> operations;
    private List<TechnologyProcessOperation> technologyProcessOperations;
    private List<TechnologyProcess> technologyProcesses;
    private List<CompatibilityOperationMachine> compatibilityOperationMachines;
    private List<Products> products;
}
