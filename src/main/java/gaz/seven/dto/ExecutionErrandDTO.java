package gaz.seven.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * Отчет Подрядчика об исполнении поручений
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExecutionErrandDTO {

    private String city;                // название города
    private Date dateRequest;           // дата составления запроса
    private String periodRequest;       // отчетный период

    private String contractorPost;      // подрядчик -> должность
    private String contractorFIO;       // подрядчик -> Ф.И.О
    private String contractorProxy;     // подрядчик -> основание (доверенность или иные документы)

    private String contractDate;        // дата договора технологических нужд ПАО «Газпром»
    private String contractNum;         // номер договора технологических нужд ПАО «Газпром»

    private String costWork;            // стоимость работ по добыче глинистых пород
    private String costWorkNDS;         // НДС от стоимости

    private String priceLicense;        // сумма расходов на приобретение лицензии на пользование недрами
    private String priceLicenseNDS;     // НДС от стоимости
    private String priceLicenseTotal;   // стоимость + НДС

}
