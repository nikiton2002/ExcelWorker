package ru.nikiton;

import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

    /*
     workingTimeMap structure:

     serviceName1 ->
                     orderNum1 ->
                                   partName1 -> workingTime
                                   partName2 -> workingTime
                     orderNum2 ->
                                   partName1 -> workingTime
                                   partName2 -> workingTime
     serviceName2 ->
                     orderNum2 ->
                                   partName3 -> workingTime
                                   partName4 -> workingTime
                     orderNum3 ->
                                   partName1 -> workingTime
                                   partName2 -> workingTime


NEW STRUCTURE:

     serviceName1 ->
                     orderNum1 ->
                                   partName1 -> 
                                                mtp1 -> workingTime
                                                mtp2 -> workingTime
                                   partName2 -> 
                                                mtp2 -> workingTime

     serviceName2 ->
                     orderNum1 ->
                                   partName3 -> 
                                                mtp1 -> workingTime
                                                mtp2 -> workingTime
                                   partName4 -> 
                                                mtp1 -> workingTime
                     orderNum2 ->
                                   partName1 -> 
                                                mtp4 -> workingTime
                                   partName2 -> 
                                                mtp1 -> workingTime
    */

public class Contractor{
    private String name;
    private int sheetIndex;  //Index of operators working time Sheet in .xls file
    private int columnIndex;     //Column number in main Sheet
    private int mainSheetIndex; // Index of contractors sheet with his worktime table
    private HashMap<String, Double> hourPrice = new HashMap<String, Double>();
    //private TreeMap<String, TreeMap<String, TreeMap<String, Double>>> workingTimeMap = new TreeMap<String, TreeMap<String, TreeMap<String, Double>>>();
    private TreeMap<String, TreeMap<String, TreeMap<String, TreeMap<String, Double>>>> workingTimeMap = new TreeMap<String, TreeMap<String, TreeMap<String, TreeMap<String, Double>>>>();

    public Contractor(String name, int sheetIndex, int mainSheetIndex) {
        this.name = name;
        this.sheetIndex = sheetIndex;
        this.columnIndex = -1;
        this.mainSheetIndex = mainSheetIndex;
    }

/*    public TreeMap<String, TreeMap<String, TreeMap<String, Double>>> getWorkingTimeMap() {//returns full copy of contractor`s Working Time Map
        TreeMap<String, TreeMap<String, TreeMap<String, Double>>> result = new TreeMap<String, TreeMap<String, TreeMap<String, Double>>>();
        for(Map.Entry<String, TreeMap<String, TreeMap<String, Double>>> serviceData : workingTimeMap.entrySet()) {
            TreeMap<String, TreeMap<String, Double>> newOrdersMap = new TreeMap<String, TreeMap<String, Double>>();
            for(Map.Entry<String, TreeMap<String, Double>> orderData : serviceData.getValue().entrySet()) {
                TreeMap<String, Double> newPartsMap = new TreeMap<String, Double>();
                for(Map.Entry<String, Double> partData : orderData.getValue().entrySet()) {
                    newPartsMap.put(partData.getKey(), partData.getValue());
                }
                newOrdersMap.put(orderData.getKey(), newPartsMap);
            }
            result.put(serviceData.getKey(), newOrdersMap);
        }
        return result;
    }*/
    public TreeMap<String, TreeMap<String, TreeMap<String, TreeMap<String, Double>>>> getWorkingTimeMap() {
        return workingTimeMap;
    }

/*    public boolean updateWorkingTimeMap(String serviceName, String orderNum, String partName, Double workingTime) {
        if("".equals(serviceName) || serviceName == null || "".equals(orderNum) || orderNum == null || partName == null || workingTime < 0) return false;
        if(workingTimeMap.containsKey(serviceName)) {
            if(workingTimeMap.get(serviceName).containsKey(orderNum)) {
                TreeMap<String, Double> orderPartsMap = workingTimeMap.get(serviceName).get(orderNum);
                if(orderPartsMap.containsKey(partName)) {
                    orderPartsMap.put(partName, orderPartsMap.get(partName) + workingTime);
                }
                else{
                    orderPartsMap.put(partName, workingTime);
                }
            }
            else{
                TreeMap<String, Double> newPartMap = new TreeMap<String, Double>();
                newPartMap.put(partName, workingTime);
                workingTimeMap.get(serviceName).put(orderNum, newPartMap);
            }
        }
        else{
            TreeMap<String, Double> newPartMap = new TreeMap<String, Double>();
            newPartMap.put(partName, workingTime);
            TreeMap<String, TreeMap<String, Double>> newOrderMap = new TreeMap<String, TreeMap<String, Double>>();
            newOrderMap.put(orderNum, newPartMap);
            workingTimeMap.put(serviceName, newOrderMap);
        }
        return true;
    }*/

    public boolean updateWorkingTimeMap(String serviceName, String orderNum, String partName, String mtp, Double workingTime) {
        if("".equals(serviceName) || serviceName == null ||
           "".equals(orderNum) || orderNum == null ||
           mtp == null || partName == null || workingTime < 0)
            return false;

        TreeMap<String, TreeMap<String, TreeMap<String, Double>>> serviceData = workingTimeMap.get(serviceName);
        if(serviceData != null) {
            TreeMap<String, TreeMap<String, Double>> orderData = serviceData.get(orderNum);
            if(orderData != null) {
                TreeMap<String, Double> partMap = orderData.get(partName);
                if(partMap != null) {
                    Double mtpValue = partMap.get(mtp);
                    if(mtpValue != null) {
                        partMap.put(mtp, mtpValue + workingTime);
                    }
                    else {
                        partMap.put(mtp, workingTime);
                    }
                }
                else {
                    partMap = new TreeMap<String, Double>();
                    partMap.put(mtp, workingTime);
                    orderData.put(partName, partMap);
                }
            }
            else {
                TreeMap<String, Double> newPartMap = new TreeMap<String, Double>();
                newPartMap.put(mtp, workingTime);
                orderData = new TreeMap<String, TreeMap<String, Double>>();
                orderData.put(partName, newPartMap);
                serviceData.put(orderNum, orderData);
            }
        }
        else {
            TreeMap<String, Double> newPartMap = new TreeMap<String, Double>();
            newPartMap.put(mtp, workingTime);
            TreeMap<String, TreeMap<String, Double>> newOrderData = new TreeMap<String, TreeMap<String, Double>>();
            newOrderData.put(partName, newPartMap);
            serviceData = new TreeMap<String, TreeMap<String, TreeMap<String, Double>>>();
            serviceData.put(orderNum, newOrderData);
            workingTimeMap.put(serviceName, serviceData);
        }
        return true;
    }

    public String getName() {
        return this.name;
    }

    public int getMainSheetIndex() {
        return this.mainSheetIndex;
    }

    public Double getHourPrice(String serviceName) {
        return hourPrice.get(serviceName);
    }

    public void setHourPrice(String serviceName, Double price) {
        if(price >= 0)
            hourPrice.put(serviceName, price);
    }

    public int getColumnIndex() {
        return this.columnIndex;
    }

    public void setColumnIndex(int position) {
        this.columnIndex = position;
    }

    public int getSheetIndex() {
        return this.sheetIndex;
    }
}