package gaz.thirteen.service;


import gaz.thirteen.dto.ConsolidatedDTO;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class ConsolidatedDataService {

   public ConsolidatedDTO getConsolidatedDTO(ConsolidatedDTO dto){

// заглушка с датой
      Date now = new Date();
      Calendar calendar = Calendar.getInstance();
      calendar.set(Calendar.DATE, calendar.getActualMinimum(Calendar.DATE));
      Date from = calendar.getTime();
      calendar.set(Calendar.DATE, calendar.getActualMaximum(Calendar.DATE));
      Date to = calendar.getTime();

      SimpleDateFormat df = new SimpleDateFormat("dd.MM.yyyy г.");

      dto.setAgent("ООО \"Газпром инвест\", 196210, город Санкт-Петербург, улица Стартовая, дом 6, литер Д. , тел.(812)4551700, факс (812)4551741");
      dto.setNumDoc("55");
      dto.setDatePreparation(df.format(now));
      dto.setPeriodFrom(df.format(from));
      dto.setPeriodTo(df.format(to));

      List<String> data = new ArrayList<>();
      List<String> total = new ArrayList<>();

      for (int i = 0; i < 22; i++) {
         data.add("d");
         if(i<5){
            total.add("t");
         }
      }

      dto.setConstructionT(total);

      dto.setQuarry(data);
      dto.setQuarryT(total);

      dto.setSubject(data);
      dto.setSubjectT(total);

      dto.setProvider(data);
      dto.setProviderT(total);

      dto.setContractor(data);
      dto.setContractorT(total);

      dto.setObject(data);
      dto.setObjectT(total);

      dto.setConstruction(data);

      return dto;
   }
}
