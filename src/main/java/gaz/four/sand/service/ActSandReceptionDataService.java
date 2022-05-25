package gaz.four.sand.service;

import gaz.four.sand.dto.ActSandReceptionDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class ActSandReceptionDataService {

    public ActSandReceptionDTO getActSandReceptionDTO (){


        // заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();

        // формат даты
        SimpleDateFormat dfFull = new SimpleDateFormat("dd MMMM yyyy г.");

        ActSandReceptionDTO dto = new ActSandReceptionDTO();

        dto.setNumQuarry("Карьер №2");
        dto.setNameQuarry("Месторождение каких-то там пород");
        dto.setDateRequest(now);
        dto.setCity("Лабытнанги");
        dto.setFirstPost("начальника отдела 647/2/4 Управления 647/2 Департамента 647 ПАО «Газпром»");
        dto.setFirstFIO("Щеткин Евгений Сергеевич");
        dto.setFirstProxy("доверенности от 09 ноября 2020 года № 01/04/04-564д");
        dto.setSecondPost("заместителя начальника управления исполнения договоров подготовки производства "
                + "ООО «Газпром инвест»");
        dto.setSecondFIO("Черепанов Дмитрий Андреевич");
        dto.setSecondProxy("доверенности от 09 ноября 2020 года № 01/04/04-564д");
        dto.setPeriodRequest("c " + dfFull.format(from) + " по " + dfFull.format(to));
        dto.setContractDate("10.04.2020г");
        dto.setContractNum("33432");
        dto.setCapacitySand("113");
        dto.setSandGOST("80568-2014");
        dto.setCodeOK("142129");
        dto.setCodeOPI("10008");
        dto.setSheetsNum("10");

        return dto;
    }
}
