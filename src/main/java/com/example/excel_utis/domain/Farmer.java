package com.example.excel_utis.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Farmer {
    private String farmerName;

    private String certType;

    private String certCode;
}
