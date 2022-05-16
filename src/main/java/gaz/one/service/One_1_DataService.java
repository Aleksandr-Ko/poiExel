package gaz.one.service;

import gaz.one.dto.One_1_DTO;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

public class One_1_DataService {

    private List<String> quarry (){

        List<String> ary = new ArrayList<>();
        ary.add("one quarry");
        ary.add("two quarry");
        ary.add("tree quarry");

        return ary;
    }

    public One_1_DTO create() {

        Date from = new Date();

        Calendar calendar = Calendar.getInstance();
        calendar.setTime(from);
        calendar.set(Calendar.DATE, 2);
        Date to = calendar.getTime();

        SimpleDateFormat df = new SimpleDateFormat("yyyy.MM");

        One_1_DTO one = new One_1_DTO();

        one.setInvestmentProject("new InvestmentProject");
        one.setOKS("new OKS");
        one.setSMR("new SMR");
        one.setFIO("Никитин Илья Платонович");
        one.setIO(printSecondName(one.getFIO()));
        one.setOPI("new OPI");
        one.setQuarry(printQuarry(quarry()));
        one.setDateRequest("с " + df.format(from) + " по " + df.format(to));

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
