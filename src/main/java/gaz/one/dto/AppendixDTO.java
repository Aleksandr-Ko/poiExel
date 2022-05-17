package gaz.one.dto;


import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class AppendixDTO {

    private String numberMail;
    private String dateRequest;
    private String datePeriod;

    private String viewOPI;
    private String nameQuarry;
    private String numberContractOPI;
    private String dateContractOPI;
    private String contractorSMR;
    private String deliveryObject;
    private String capacity;
    private String FIO;

}
