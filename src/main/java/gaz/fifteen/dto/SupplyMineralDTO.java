package gaz.fifteen.dto;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * ЕЖЕДНЕВНАЯ СВОДКА О ПОСТАВКЕ ОПИ ДЛЯ ОКС
 */

@Data
public class SupplyMineralDTO {

    private String nameInvestProject;           // наименование инвестпроекта
    private String codeInvestProject;           // код инвестпроекта

    List<String> dummy = new ArrayList<>();     // болванка для данных

}
