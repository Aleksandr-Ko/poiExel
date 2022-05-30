package gaz.five.clay.service;

import gaz.five.clay.dto.ActClayDoneWorkDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class ActClayDoneWorkDataService {

    public ActClayDoneWorkDTO getActClayDoneWorkDTO () {

// заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();

// формат даты
        SimpleDateFormat dfFull = new SimpleDateFormat("dd MMMM yyyy г.");

        ActClayDoneWorkDTO dto = new ActClayDoneWorkDTO();

        dto.setNumQuarry("2");
        dto.setNameQuarry("Месторождение каких-то там пород");
        dto.setDateRequest(now);
        dto.setCity("Санкт-Петербург");
        dto.setFirstPost("начальника отдела 647/2/4 Управления 647/2 Департамента 647 ПАО «Газпром»");
        dto.setFirstFIO("Щеткин Евгений Сергеевич");
        dto.setFirstProxy("доверенности от 09 ноября 2020 года № 01/04/04-564д");
        dto.setSecondPost("заместителя начальника управления исполнения договоров подготовки производства "
                + "ООО «Газпром инвест»");
        dto.setSecondFIO("Черепанов Дмитрий Андреевич");
        dto.setSecondProxy("доверенности от 15 июля 2020 года № 01/8158");
        dto.setPeriodRequest("c " + dfFull.format(from) + " по " + dfFull.format(to));
        dto.setContractDate("04 декабря 2017");
        dto.setContractNum("ОПИ-51");
        dto.setLicense("СЛХ 81417 ТЭ");
        dto.setCostWork("39241357.76");
        dto.setCostWorkNDS("7848271.55");
        dto.setCostWorkTotal("47089629.31");
        dto.setNumPaymentOrder("-");
        dto.setDatePaymentOrder("-");
        dto.setPrepayPercentage("-");
        dto.setPrepay("-");
        dto.setPrepayNDS("-");
        dto.setDeferPay("2354481.47");
        dto.setPayment("44735147.84");
        dto.setPaymentNDS("7455857.97");
        dto.setViewClay("Глинистых пород");
        dto.setCapacityClay("113");
        dto.setClayGOST("80568-2014");
        dto.setCodeOK("142129");
        dto.setCodeOPI("10008");

        return dto;
    }
}
