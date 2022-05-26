package gaz.four.clay.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ActClayReceptionDTO {

    private String city;                // название города
    private Date dateRequest;           // дата составления запроса
    private String periodRequest;       // отчетный период

    private String nameQuarry;          // наименование карьера
    private String numQuarry;           // номер карьера

    private String firstPost;           // заказчик -> должность
    private String firstFIO;            // заказчик -> Ф.И.О
    private String firstProxy;          // заказчик -> основание (доверенность или иные документы)

    private String secondPost;          // генподрядчик -> должность
    private String secondFIO;           // генподрядчик -> Ф.И.О
    private String secondProxy;         // генподрядчик -> основание (доверенность или иные документы)

    private String contractDate;        // дата договора генерального подряда
    private String contractNum;         // номер договора генерального подряда

    private String viewClay;            // вид глинистых грунтов
    private String capacityClay;        // объем глинистых грунтов

    private String clayGOST;            // стандарт ГОСТ глинистых грунтов
    private String codeOK;              // код по ОК
    private String codeOPI;             // код вида ОПИ
    private String sheetsNum;           // кол-во листов накладных
}
