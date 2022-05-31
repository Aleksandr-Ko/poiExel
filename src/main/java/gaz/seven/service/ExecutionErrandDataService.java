package gaz.seven.service;

import gaz.seven.dto.ExecutionErrandDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class ExecutionErrandDataService {

    public ExecutionErrandDTO getExecutionErrandDTO() {
// заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();
// формат даты
        SimpleDateFormat dfFull = new SimpleDateFormat("dd MMMM yyyy г.");
        ExecutionErrandDTO dto = new ExecutionErrandDTO();

        dto.setCity("Санкт-Петербург");
        dto.setPeriodRequest("c " + dfFull.format(from) + " по " + dfFull.format(to));

        dto.setContractorPost("заместителя начальника управления исполнения договоров подготовки производства "
                + "ООО «Газпром инвест»");
        dto.setContractorFIO("Черепанов Дмитрий Андреевич");
        dto.setContractorProxy("доверенности от 15 июля 2020 года № 01/8158");

        dto.setContractDate("04 декабря 2017");
        dto.setContractNum("ОПИ-51");

        dto.setCostWorkNDS("7848271.55");
        dto.setCostWork("39241357.76");

        dto.setPriceLicenseTotal("-");
        dto.setPriceLicenseNDS("-");
        dto.setPriceLicense("-");

// TODO ->  какую дату выводить? (запрос, составления или отчетного периода)
        dto.setDateRequest(now);

        return dto;
    }
}
