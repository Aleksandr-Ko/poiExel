package gaz.two.service;

import gaz.two.dto.ActAcquisitionOPIDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class ActAcquisitionOPIDataService {

    public static void main(String[] args) {

    }

    public ActAcquisitionOPIDTO createDTO() {
        // заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();

        // формат даты
        SimpleDateFormat dfFull = new SimpleDateFormat("dd MMMM yyyy");
        SimpleDateFormat dfquotes = new SimpleDateFormat("«dd» MMMM yyyy г.");

        ActAcquisitionOPIDTO act = new ActAcquisitionOPIDTO();

        act.setCity("Санкт-Петербург");
        act.setDateRequest("" + dfquotes.format(now));
        act.setPeriodRequest("с "+ dfFull.format(from) + " по " + dfFull.format(to));
        act.setContractNum("ОПИ-51");
        act.setContractDate("«04» декабря 2017 г");
        act.setNameOPI("Песок (ГОСТ 8736-2014) 800000052 ООО «Регион-Сбыт»");
        act.setFirstPost("начальника отдела 647/2/4 Управления 647/2 Департамента 647 ПАО «Газпром»");
        act.setFirstFIO("Щеткин Евгений Сергеевич");
        act.setFirstProxy("доверенности от 09 ноября 2020 года № 01/04/04-564д");
        act.setSecondPost("заместителя начальника управления исполнения договоров подготовки производства "
                + "ООО «Газпром инвест»");
        act.setSecondFIO("Черепанова Дмитрия Андреевича");
        act.setSecondProxy("доверенности от 09 ноября 2020 года № 01/04/04-564д");

        act.setExpenses("3 861 380 (Три миллиона восемьсот шестьдесят одна тысяча триста восемьдесят) рублей "
                + "05 копеек");
        act.setExpensesNds("643 563 (Шестьсот сорок три тысячи пятьсот шестьдесят три) рубля "
                + "34 копейки");
        act.setReward("4 260 (Четыре "
                + "тысячи двести шестьдесят) рублей 19 копеек");
        act.setRewardNDS("852 (Восемьсот пятьдесят два) рубля "
                + "04 копейки");
        act.setRewardTotal("5 112 (Пять тысяч сто двенадцать) рублей 23 копейки");
        act.setPrice("4260,19");
        act.setPriceNDS("852,04");
        act.setPriceTotal("5112,23");

        return act;
    }
}
