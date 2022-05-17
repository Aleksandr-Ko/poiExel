package gaz.one.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class RequestToSupplierOPIDTO {

    private String investmentProject;
    private String OKS;
    private String SMR;
    private String FIO;
    private String IO;
    private String OPI;

    private String datePeriod;
    private String quarry;



}
