package gaz.one.service;

import gaz.one.dto.RequestToSupplierOPIDTO;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

public class RequestToSupplierOPIDataService {

    private List<String> quarry (){

        List<String> listquarry = new ArrayList<>();
        listquarry.add("«one quarry»");
        listquarry.add("«two quarry»");
        listquarry.add("«tree quarry»");

        return listquarry;
    }

    public RequestToSupplierOPIDTO create() {

        Date from = new Date();

        Calendar calendar = Calendar.getInstance();
        calendar.setTime(from);
        calendar.add(Calendar.MONTH, 2);
        Date to = calendar.getTime();

        SimpleDateFormat df = new SimpleDateFormat("yyyy.MM");

        RequestToSupplierOPIDTO one = new RequestToSupplierOPIDTO();

        one.setInvestmentProject("new InvestmentProject");
        one.setOKS("new OKS");
        one.setSMR("new SMR");
        one.setFIO("Никитин Илья Платонович");
        one.setIO(printSecondName(one.getFIO()));
        one.setOPI("new OPI");
        one.setQuarry(printQuarry(quarry()));
        one.setDatePeriod("с " + df.format(from) + " по " + df.format(to));

        return one;
    }

    private String printSecondName(String fullName) {
        String[] word = fullName.split(" ");
        StringBuilder nameIO = new StringBuilder();
        // вывод выводит имя и отчество
        for (int i = 0; i < word.length; i++) {
            if (i != 0) {
                nameIO.append(word[i]).append(" ");
            }
        }
        return nameIO.toString();
    }

    private String printQuarry(List<String> quarry) {
        return quarry.stream()
                .map(String::valueOf)
                .collect(Collectors.joining(", "));
    }
}
