package gaz.thirteen.dto;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class ConsolidatedDTO {

    private String Agent;                               // полное наименование агента (имя, адрес, телефон)

    private String numDoc;                              // номер документа
    private String datePreparation;                     // дата составления
    private String periodFrom;                          // отчетный период с
    private String periodTo;                            // отчетный период по

    List<String> construction = new ArrayList<>();      // болванка карьеров
    List<String> constructionT = new ArrayList<>();     // болванка карьеров

    List<String> quarry = new ArrayList<>();            // болванка карьеров
    List<String> quarryT = new ArrayList<>();           // болванка карьеров

    List<String> subject = new ArrayList<>();           // болванка субъектов
    List<String> subjectT = new ArrayList<>();          // болванка субъектов

    List<String> provider = new ArrayList<>();          // болванка поставщиков
    List<String> providerT = new ArrayList<>();         // болванка поставщиков

    List<String> contractor = new ArrayList<>();        // болванка подрядчиков
    List<String> contractorT = new ArrayList<>();       // болванка подрядчиков

    List<String> object = new ArrayList<>();            // болванка объектов
    List<String> objectT = new ArrayList<>();           // болванка объектов

}
