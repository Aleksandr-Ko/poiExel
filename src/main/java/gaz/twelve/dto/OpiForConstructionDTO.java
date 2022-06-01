package gaz.twelve.dto;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * СПРАВКИ О ДВИЖЕНИИ ОПИ, ПРЕДНАЗНАЧЕННЫХ ДЛЯ СТРОИТЕЛЬСТВА
 */
@Data
public class OpiForConstructionDTO {

    private String customer;                    // заказчик
    private String contractor;                  // генподрядчик
    private String investProject;               // наименование инвестиционного проекта
    private String codeConstr;                  // код стройки
    private String nameObjConstr;               // наименование объекта
    private String codeObj;                     // код объекта
    private String generalContract;             // договор генподряда

    private String numDoc;                      // номер документа
    private String datePreparation;             // дата составления
    private String periodFrom;                  // отчетный период с
    private String periodТо;                    // отчетный период по

    List<String> dummy = new ArrayList<>();     // болванка для данных
    List<String> total = new ArrayList<>();     // болванка для итоговых значений

}
