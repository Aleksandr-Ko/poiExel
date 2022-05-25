package gaz.three.crag.service;

import gaz.three.crag.dto.ActCragAcceptanceDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class ActCragAcceptanceDataService {


    public ActCragAcceptanceDTO getActCragAcceptanceDTO() {
        // заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();

        // формат даты
        SimpleDateFormat dfFull = new SimpleDateFormat("dd MMMM yyyy г.");

        ActCragAcceptanceDTO dto = new ActCragAcceptanceDTO();

        dto.setNumQuarry("283-1М");
        dto.setDateRequest(now);
        dto.setCity("Санкт-Петербург");
        dto.setFirstPost("начальника отдела 647/2/4 Управления 647/2 Департамента 647 ПАО «Газпром»");
        dto.setFirstFIO("Щеткин Евгений Сергеевич");
        dto.setFirstProxy("доверенности от 09 ноября 2020 года № 01/04/04-564д");
        dto.setSecondPost("заместителя начальника управления исполнения договоров подготовки производства "
                + "ООО «Газпром инвест»");
        dto.setSecondFIO("Черепанов Дмитрий Андреевич");
        dto.setSecondProxy("доверенности от 09 ноября 2020 года № 01/04/04-564д");
        dto.setContractNum("ОПИ-51");
        dto.setContractDate("«04» декабря 2017");
        dto.setPeriodRequest("с "+ dfFull.format(from) + " по " + dfFull.format(to));

        dto.setViewCrag("Песок (ГОСТ 8736-2014) 800000052 ООО «Регион-Сбыт»");

        dto.setNameQuarry("«Пруды-Моховое-Яскинское» АО \"ЛСР. Базовые материалы\"");
        dto.setCapacityCrag("113");
        dto.setStandartGOST("80568-2014");
        dto.setCodeOK("142129");
        dto.setCodeOPI("10008");
        dto.setSheetsNum("10");

        return  dto;
    }
}
