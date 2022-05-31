package gaz.ten.service;

import gaz.ten.dto.ZGAgentExpenseDTO;

import java.util.List;

public class ZGAgentExpenseDataService {

    public ZGAgentExpenseDTO getZGAgentExpenseDTO (ZGAgentExpenseDTO dto){

        List<String> total = dto.getTotal();

        for (int i = 0; i < 20; i++) {
            total.add("1");
        }

        return dto;
    }
}
