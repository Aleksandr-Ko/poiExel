package gaz.six.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * Отчет Подрядчика о приобретении ОПИ (Материалов)
 * и контроле их качества
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ObtainOPIDTO {


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
    private String costWorkTotal;       // стоимость + НДС

    private String nameCounterparty;    // Наименование контрагента
    private String numDateContr;        // Номер и дата договора/ соглашения
    private String termContr;           // Срок действия договора/ соглашения
    private String subjectContr;        // Предмет договора, существенные условия договора/ соглашения

}
