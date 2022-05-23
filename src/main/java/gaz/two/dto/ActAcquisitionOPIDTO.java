package gaz.two.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ActAcquisitionOPIDTO {

    private String city;
    private Date dateRequest;
    private String periodRequest;

    private String contractNum;
    private String contractDate;

    private String nameOPI;

    private String firstPost;
    private String firstFIO;
    private String firstProxy;

    private String secondPost;
    private String secondFIO;
    private String secondProxy;

    private String expenses;
    private String expensesNds;

    private String reward;
    private String rewardNDS;
    private String rewardTotal;

    private String price;
    private String priceNDS;
    private String priceTotal;



}
