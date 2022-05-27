package gaz.five.sand.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * АКТ
 * выполненных работ по добыче песка
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ActSandDoneWorkDTO {

    private String numQuarry;
    private String nameQuarry;
    private Date dateRequest;
    private String periodRequest;
    private String city;
    private String firstPost;
    private String firstFIO;
    private String firstProxy;
    private String secondPost;
    private String secondFIO;
    private String secondProxy;
    private String contractDate;
    private String contractNum;
    private String capacitySand;
    private String sandGOST;
    private String codeOK;
    private String codeOPI;
    private String license;
    private String costWork;
    private String costWorkNDS;
    private String costWorkTotal;
    private String numPaymentOrder;
    private String datePaymentOrder;
    private String prepayPercentage;
    private String prepay;
    private String prepayNDS;
    private String deferPay;
    private String payment;
    private String paymentNDS;

}
