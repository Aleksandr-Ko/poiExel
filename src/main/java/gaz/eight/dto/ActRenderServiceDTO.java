package gaz.eight.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * АКТ
 * сдачи-приемки оказанных услуг Подрядчика
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ActRenderServiceDTO {

    private String city;                // название города
    private Date dateRequest;           // дата составления запроса
    private String periodRequest;       // отчетный период

    private String firstPost;           // заказчик -> должность
    private String firstFIO;            // заказчик -> Ф.И.О
    private String firstProxy;          // заказчик -> основание (доверенность или иные документы)

    private String secondPost;          // подрядчик -> должность
    private String secondFIO;           // подрядчик -> Ф.И.О
    private String secondProxy;         // подрядчик -> основание (доверенность или иные документы)

    private String contractDate;        // дата договора технологических нужд ПАО «Газпром»
    private String contractNum;         // номер договора технологических нужд ПАО «Газпром»

    private String costWork;            // стоимость работ
    private String costWorkNDS;         // НДС от стоимости

    private String priceLicense;        // сумма расходов на приобретение лицензии на пользование недрами
    private String priceLicenseNDS;     // НДС от стоимости
    private String priceLicenseTotal;   // стоимость + НДС

    private String reward;              // вознаграждение
    private String rewardNDS;           // НДС от вознаграждения
    private String rewardTotal;         // вознаграждение + НДС


}
