package gaz.eight.service;

import gaz.eight.dto.ActRenderServiceDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class ActRenderServiceDataService {

    public ActRenderServiceDTO getActRenderServiceDTO (){



        // заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();
// формат даты
        SimpleDateFormat dfFull = new SimpleDateFormat("dd MMMM yyyy г.");

        ActRenderServiceDTO dto = new ActRenderServiceDTO();

        dto.setCity("Санкт-Петербург");
        dto.setPeriodRequest("c " + dfFull.format(from) + " по " + dfFull.format(to));

        dto.setFirstPost("начальника отдела 647/2/4 Управления 647/2 Департамента 647 ПАО «Газпром»");
        dto.setFirstFIO("Щеткин Евгений Сергеевич");
        dto.setFirstProxy("доверенности от 09 ноября 2020 года № 01/04/04-564д");
        dto.setSecondPost("заместителя начальника управления исполнения договоров подготовки производства "
                + "ООО «Газпром инвест»");
        dto.setSecondFIO("Черепанов Дмитрий Андреевич");
        dto.setSecondProxy("доверенности от 15 июля 2020 года № 01/8158");
        dto.setContractDate("04 декабря 2017");
        dto.setContractNum("ОПИ-51");
        dto.setCostWorkNDS("7848271.55");
        dto.setCostWork("39241357.76");
        dto.setPriceLicenseTotal("-");
        dto.setPriceLicenseNDS("-");
        dto.setPriceLicense("-");
        dto.setReward("4260");
        dto.setRewardNDS("852");
        dto.setRewardTotal("5112");

// TODO ->  какую дату выводить? (запрос, составления или отчетного периода)
        dto.setDateRequest(now);

        return dto;
    }
}
