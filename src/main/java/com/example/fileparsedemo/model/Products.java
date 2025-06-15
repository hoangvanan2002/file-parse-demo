package com.example.fileparsedemo.model;

import lombok.*;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Products {
    private String productCode;
    private String customerCode;
    private String productName;
    private String productEnName;
    private String productLine;
    private String productType;
    private String deliveryCharacteristicCode;
    private String productModel;
    private String unit;
}
