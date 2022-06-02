package gaz.fourteen.dto;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * ЕЖЕДНЕВНАЯ СВОДКА РЕЗУЛЬТАТОВ РАБОТ НА КАРЬЕРЕ
 */
@Data
public class QuarryWorkDTO {

    private String nameInvestProject;           // наименование инвестпроекта
    private String codeInvestProject;           // код инвестпроекта

    List<String> dummy = new ArrayList<>();     // болванка для данных
}
