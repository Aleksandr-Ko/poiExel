package gaz.five.sand.service;

import gaz.five.sand.dto.ActSandDoneWorkDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class ActSandDoneWorkDataService {

    public ActSandDoneWorkDTO getActSandDoneWorkDTO() {

// заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();

// формат даты
        SimpleDateFormat dfFull = new SimpleDateFormat("dd MMMM yyyy г.");

        ActSandDoneWorkDTO dto = new ActSandDoneWorkDTO();

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
        dto.setCapacitySand("113");
        dto.setSandGOST("80568-2014");
        dto.setCodeOK("142129");
        dto.setCodeOPI("10008");
        dto.setLicense("СЛХ 81417 ТЭ");
        dto.setCostWork("39 241 357 (Тридцать девять миллионов двести сорок одна тысяча триста пятьдесят семь"
                + " рублей 76коп.");
        dto.setCostWorkNDS("7 848 271 (семь миллионов восемьсот сорок восемь тысяч двести семьдесят один"
        + ") рубль 55 коп.");
        dto.setCostWorkTotal("47 089 629 (сорок семь миллионов восемьдесят девять тысяч шестьсот двадцать девять"
        + ") рублей 31 коп.");
        dto.setNumPaymentOrder("-");
        dto.setDatePaymentOrder("-");
        dto.setPrepayPercentage("-");
        dto.setPrepay("-");
        dto.setPrepayNDS("-");
        dto.setDeferPay("2 354 481 (Два миллиона триста пятьдесят четыре тысячи четыреста восемьдесят один"
        + ") рубль 47 коп");
        dto.setPayment("44 735 147 (сорок четыре миллиона семьсот тридцать пять тысяч сто сорок семь)"
        + " рублей 84 коп.");
        dto.setPaymentNDS("7 455 857 (Семь миллионов четыреста пятьдесят пять тысяч восемьсот пятьдесят семь"
        + ") рублей 97 коп.");

        return dto;

    }
}
