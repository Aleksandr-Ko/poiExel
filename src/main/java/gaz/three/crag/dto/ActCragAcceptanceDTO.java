package gaz.three.crag.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ActCragAcceptanceDTO {

    private String city;                // название города
    private Date dateRequest;           // дата составления запроса
    private String periodRequest;       // отчетный период

    private String contractNum;         // номер договора ПАО «Газпром» на добычу ОПИ
    private String contractDate;        // дата договора ПАО «Газпром» на добычу ОПИ

    private String numQuarry;           // номер карьера
    private String nameQuarry;          // наименование карьера

    private String firstPost;           // заказчик -> должность
    private String firstFIO;            // заказчик -> Ф.И.О
    private String firstProxy;          // заказчик -> основание (доверенность или иные документы)

    private String secondPost;          // подрядчик -> должность
    private String secondFIO;           // подрядчик -> Ф.И.О
    private String secondProxy;         // подрядчик -> основание (доверенность или иные документы)

    private String viewCrag;            // вид скальных пород
    private String capacityCrag;        // объем скальных пород

    private String standartGOST;        // стандарт ГОСТ скальных пород
    private String codeOK;              // код по ОК
    private String codeOPI;             // код вида ОПИ
    private String sheetsNum;           // кол-во листов накладных
}
