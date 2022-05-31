package gaz.nine.service;

import gaz.nine.dto.ZVAgentExpenseDTO;

import java.util.List;


public class ZVAgentExpenseDataService {

    public ZVAgentExpenseDTO getZVAgentExpenseDTO (ZVAgentExpenseDTO dto){

        List<String> dummy = dto.getDummy();
        List<String> total = dto.getTotal();

        for (int i = 0; i < 18; i++) {
            dummy.add("1");
            if (i > 1) {
                total.add("tot");
            }
        }
        return dto;
    }
}
