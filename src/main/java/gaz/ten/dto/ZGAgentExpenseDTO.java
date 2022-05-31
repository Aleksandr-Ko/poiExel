package gaz.ten.dto;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * 3Г ОТЧЕТА АГЕНТА
 * «О РАСХОДЕ ОСНОВНЫХ МАТЕРИАЛОВ В СТРОИТЕЛЬСТВЕ В СОПОСТАВЛЕНИИ С РАСХОДОМ, ОПРЕДЕЛЕННЫМ ПО ПРОИЗВОДСТВЕННЫМ НОРМАМ»
 */
@Data
public class ZGAgentExpenseDTO {
    List<String> total = new ArrayList<>();
}
