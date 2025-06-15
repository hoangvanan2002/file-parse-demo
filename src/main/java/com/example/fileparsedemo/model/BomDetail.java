package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BomDetail {
    private String productCode;
    private String bomLevel;
    private String itemCode;
    private String type;
    private String quantity;
    private String componentYield;
    private String technologyProcessCode;
}
