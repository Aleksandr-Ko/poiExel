package gaz.one.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class NeedsOPIDTO {


    private String contractorSMR;       // Подрядчик СМР

    private String datePeriod;          // Отчетный период поставки (год, месяц)
    private String dateRequest;         // дата формирования заявки

    private String investmentProject;   // Наименование инвестпроекта

    private String nameOKS;             // Наименование ОКС

    private String listNameQuarry;      // список карьеров для первой страницы
    private String nameQuarry;          // Наименование карьера
    private String supplierOPI;         // Поставщик ОПИ
    private String viewOPI;             // Вид ОПИ
    private String numberContractOPI;   // Номер договора Поставщика ОПИ
    private String dateContractOPI;     // Дата договора Поставщика ОПИ

    private String deliveryObject;      // Объект поставки
    private String capacity;            // Объем

}
