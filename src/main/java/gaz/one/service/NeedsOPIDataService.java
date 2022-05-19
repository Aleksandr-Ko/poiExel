package gaz.one.service;

import gaz.one.dto.NeedsOPIDTO;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

public class NeedsOPIDataService {

    private List<String> quarry (){
        List<String> listquarry = new ArrayList<>();
        listquarry.add("«one quarry»");
        listquarry.add("«two quarry»");
        listquarry.add("«tree quarry»");
        return listquarry;
    }
    public NeedsOPIDTO create() {
        // заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.add(Calendar.MONTH, 2);
        Date from = calendar.getTime();
        calendar.add(Calendar.MONTH, 6);
        Date to = calendar.getTime();

        // формат даты
        SimpleDateFormat df = new SimpleDateFormat("yyyy.MM");
        SimpleDateFormat dfFull = new SimpleDateFormat("dd.MM.yyyy");

        NeedsOPIDTO one = new NeedsOPIDTO();

        one.setDateRequest(""+dfFull.format(now));
        one.setInvestmentProject("new InvestmentProject");
        one.setNameOKS("new OKS");
        one.setContractorSMR("new SMR");
        one.setSupplierOPI("new OPI");
        one.setViewOPI("sand");
        one.setListNameQuarry(printQuarry(quarry()));
        one.setDatePeriod(" " + df.format(from) + " по " + df.format(to));
        one.setNameQuarry("new name Quarry");
        one.setNumberContractOPI("new numContr");
        one.setDateContractOPI("new dateContr");
        one.setDeliveryObject("new Object");
        one.setCapacity("100");

        return one;
    }

    // метод выводит для page_1 список карьеров
    private String printQuarry(List<String> quarry) {
        return quarry.stream()
                .map(String::valueOf)
                .collect(Collectors.joining(", "));
    }
}
