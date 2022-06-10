package gaz.six.service;

import gaz.six.dto.ObtainOPIDTO;

import java.text.SimpleDateFormat;
import java.util.*;

public class ObtainOPIDataService {

    public ObtainOPIDTO getObtainOPIDTO(ObtainOPIDTO dto) {
// заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();
// формат даты
        SimpleDateFormat dfFull = new SimpleDateFormat("dd MMMM yyyy г.");

//        ObtainOPIDTO dto = new ObtainOPIDTO();

        dto.setCity("Санкт-Петербург");
        dto.setContractorPost("заместителя начальника управления исполнения договоров подготовки производства "
                + "ООО «Газпром инвест»");
        dto.setContractorFIO("Черепанов Дмитрий Андреевич");
        dto.setContractorProxy("доверенности от 15 июля 2020 года № 01/8158");

        dto.setCostWorkTotal("47089629.31");
        dto.setCostWorkNDS("7848271.55");
        dto.setCostWork("39241357.76");
        dto.setContractDate("04 декабря 2017");
        dto.setContractNum("ОПИ-51");
        dto.setPeriodRequest("c " + dfFull.format(from) + " по " + dfFull.format(to));

        List<String>ls = dto.getDummy();

        for (int i = 0; i < 9; i++) {
            ls.add("" + i);
        }


// TODO ->  какую дату выводить? (запрос, составления или отчетного периода)
        dto.setDateRequest(now);

        return dto;
    }
}
