package gaz.eleven.service;

import gaz.eleven.dto.ZDAgentExpenseDTO;

import java.util.List;

public class ZDAgentExpenseDataService {
    public ZDAgentExpenseDTO getZDAgentExpenseDTO(ZDAgentExpenseDTO dto) {

        List<String> dummy = dto.getDummy();
        List<String> total = dto.getTotal();

        for (int i = 0; i < 10; i++) {
            dummy.add("1");
            if (i > 4) {
                total.add("tot");
            }
        }

        return dto;
    }
}
