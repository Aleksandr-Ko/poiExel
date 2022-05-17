package gaz.one.service;

import gaz.one.dto.AppendixDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class AppendixDataService {

    public AppendixDTO create() {

        Date from = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(from);
        calendar.add(Calendar.MONTH, 2);
        Date to = calendar.getTime();

        SimpleDateFormat df = new SimpleDateFormat("yyyy.MM");
        SimpleDateFormat dfFull = new SimpleDateFormat("dd.MM.yyyy");

        AppendixDTO appendix = new AppendixDTO();

        appendix.setNumberMail("???");
        appendix.setDateRequest(""+dfFull.format(from));
        appendix.setDatePeriod("с " + df.format(from) + " по " + df.format(to));
        appendix.setViewOPI("new OPI");
        appendix.setNameQuarry("new Quarry");
        appendix.setNumberContractOPI("new numContr");
        appendix.setDateContractOPI("new dateContr");
        appendix.setContractorSMR("new SMR");
        appendix.setDeliveryObject("new Object");
        appendix.setCapacity("100");
        appendix.setFIO("Никитин Илья Платонович");


        return appendix;
    }
}
