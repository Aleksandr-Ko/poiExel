package gaz.five.crag.dto;


import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * АКТ
 * выполненных работ по добыче Скальных пород
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ActCragDoneWorkDTO {

    private String city;                // название города
    private Date dateRequest;           // дата составления запроса
    private String periodRequest;       // отчетный период

    private String nameQuarry;          // наименование карьера
    private String numQuarry;           // номер карьера

    private String firstPost;           // заказчик -> должность
    private String firstFIO;            // заказчик -> Ф.И.О
    private String firstProxy;          // заказчик -> основание (доверенность или иные документы)

    private String secondPost;          // подрядчик -> должность
    private String secondFIO;           // подрядчик -> Ф.И.О
    private String secondProxy;         // подрядчик -> основание (доверенность или иные документы)

    private String contractDate;        // дата договора технологических нужд ПАО «Газпром»
    private String contractNum;         // номер договора технологических нужд ПАО «Газпром»

    private String license;             // вид Лицензии, серия и номер

    private String viewCrag;            // вид скальных пород
    private String capacityCrag;        // объем скальных пород

    private String cragGOST;            // стандарт ГОСТ скальных пород
    private String codeOK;              // код по ОК
    private String codeOPI;             // код вида ОПИ

    private String costWork;            // стоимость работ по добыче Скальных пород
    private String costWorkNDS;         // НДС от стоимости
    private String costWorkTotal;       // стоимость + НДС

    private String numPaymentOrder;     // номер платежного поручения
    private String datePaymentOrder;    // дата платежного поручения
    private String prepayPercentage;    // процент аванса
    private String prepay;              // аванс
    private String prepayNDS;           // НДС от аванса

    private String deferPay;            // Отложенный платеж

    private String payment;             // полная стоимость по акту
    private String paymentNDS;          // НДС от полной стоимости

}
