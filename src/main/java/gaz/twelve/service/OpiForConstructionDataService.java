package gaz.twelve.service;

import gaz.twelve.dto.OpiForConstructionDTO;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class OpiForConstructionDataService {

    public OpiForConstructionDTO getOpiForConstructionDTO(OpiForConstructionDTO dto) {

// заглушка с датой
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
        Date from = calendar.getTime();
        calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
        Date to = calendar.getTime();

        SimpleDateFormat df = new SimpleDateFormat("dd.MM.yyyy г.");

        dto.setCustomer("ПАО «Газпром»");
        dto.setContractor("ООО «Газпром инвест»");
        dto.setInvestProject("Магистральный газопровод ");
        dto.setCodeConstr("МГ-2");
        dto.setNameObjConstr("Сила Сибири");
        dto.setCodeObj("055-2007890");
        dto.setGeneralContract("Договор на добычу ОПИ");

        dto.setNumDoc("55");
        dto.setDatePreparation(df.format(now));
        dto.setPeriodFrom(df.format(from));
        dto.setPeriodТо(df.format(to));

        List<String> dummy = dto.getDummy();
        List<String> total = dto.getTotal();

        for (int i = 0; i < 18; i++) {
            dummy.add("1");
            if (i >= 1) {
                total.add("tot");
            }
        }

        return dto;
    }
}
