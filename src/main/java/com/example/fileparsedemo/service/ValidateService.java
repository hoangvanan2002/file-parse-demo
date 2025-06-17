package com.example.fileparsedemo.service;

import com.example.fileparsedemo.model.BomDetail;
import com.example.fileparsedemo.model.Operation;

import java.util.List;

public class ValidateService {

    //Tao mot doi tuong LOi, sau do check dk o tung ham 1 xem bat buoc va khong dooc trung,
    //Neu khong co loi thi in file response
    //Neu co loi thi chi tra ve master_data voi cac dong loi duoc boi do

    public static boolean isBomDetailValid(List<BomDetail> bomDetails){

        return false;
    }

    public static boolean isOperationValid(List<Operation> operations){
        for (Operation operation : operations){
            if(operation.getOperationGroup().equals("")){
                return false;
            } else {

            }
        }
        return false;
    }
}
