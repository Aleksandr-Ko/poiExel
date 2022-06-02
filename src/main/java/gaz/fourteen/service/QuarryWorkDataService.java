package gaz.fourteen.service;

import gaz.fourteen.dto.QuarryWorkDTO;

import java.util.List;

public class QuarryWorkDataService {
    
    public QuarryWorkDTO getQuarryWorkDTO (QuarryWorkDTO dto){

        dto.setNameInvestProject("Магистральный газопровод \"Сила Сибири\"");
        dto.setCodeInvestProject("055-2007890");

        List<String> dummy = dto.getDummy();

        for (int i = 0; i < 10; i++) {
            dummy.add("data");
        }

        return dto;
    }
    
    
}
