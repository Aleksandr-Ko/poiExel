package gaz.fifteen.service;


import gaz.fifteen.dto.SupplyMineralDTO;

import java.util.List;

public class SupplyMineralDataService {

    public SupplyMineralDTO getSupplyMineralDTO(){

        SupplyMineralDTO dto = new SupplyMineralDTO();

        dto.setNameInvestProject("Магистральный газопровод \"Сила Сибири\"");
        dto.setCodeInvestProject("055-2007890");

        List<String> dummy = dto.getDummy();

        for (int i = 0; i < 14; i++) {
            dummy.add("data");
        }

        return dto;
    }
}
