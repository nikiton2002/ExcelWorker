package ru.nikiton;

import java.util.*;

public class Operator{
    private String name;
    private int sheetIndex;     //Index of operators working time Sheet in .xls file
    private int columnIndex;    //Column number in main Sheet
    private String columnLetter; //Column Letter in main Sheet
    private int percentColumnIndex;
    private String percentColumnLetter;
    private TreeMap<String, TreeMap<String, TreeMap<String, Double>>> workingTimeMap = new TreeMap<String, TreeMap<String, TreeMap<String, Double>>>();

    public Operator(String name, int sheetIndex) {
        this.name = name;
        this.sheetIndex = sheetIndex;
        this.columnIndex = -1;
        this.percentColumnIndex = -1;
    }

    public String getName() {
        return this.name;
    }

    public boolean updateWorkingTimeMap(String orderNum, String partName, Double workingTime, String mtp) {
        if(orderNum==null || "".equals(orderNum) || partName==null || workingTime < 0) return false;
        TreeMap<String, TreeMap<String, Double>> orderParts = workingTimeMap.get(orderNum);
        if(orderParts != null) {
            TreeMap<String, Double> partMTPs = orderParts.get(partName);
            if(partMTPs != null) {
                Double oldWorkTime = partMTPs.get(mtp);
                if(oldWorkTime != null) {
                    partMTPs.put(mtp, oldWorkTime + workingTime);
                }
                else {
                    partMTPs.put(mtp, workingTime);
                }
            }
            else {
                TreeMap<String, Double> newMTP = new TreeMap<String, Double>();
                newMTP.put(mtp, workingTime);
                orderParts.put(partName, newMTP);
            }
        }
        else {
            TreeMap<String, TreeMap<String, Double>> newPart = new TreeMap<String, TreeMap<String, Double>>();
            TreeMap<String, Double> newMTP = new TreeMap<String, Double>();
            newMTP.put(mtp, workingTime);
            newPart.put(partName, newMTP);
            workingTimeMap.put(orderNum, newPart);
        }
        return true;
    }

    public TreeMap<String, TreeMap<String, TreeMap<String, Double>>> getWorkingTimeMap() {
/*        TreeMap<String, TreeMap<String, TreeMap<String, Double>>> result = new TreeMap<String, TreeMap<String, TreeMap<String, Double>>>();
        for(Map.Entry<String, TreeMap<String, TreeMap<String, Double>>> orderData : workingTimeMap.entrySet()) {
            result.put(orderData.getKey(), orderData.getValue());
        }*/
        return workingTimeMap;
    }

    public void setColumnIndex(int position, String columnLetter) {
        this.columnIndex = position;
        this.columnLetter = columnLetter;
        this.percentColumnIndex = position+1;
        this.percentColumnLetter = Library.getColumnLetter(percentColumnIndex);
    }

    public int getColumnIndex() {
        return this.columnIndex;
    }

    public int getPercentColumnIndex() {
        return this.percentColumnIndex;
    }

    public String getColumnLetter() {
        return this.columnLetter;
    }

    public String getPercentColumnLetter() {
        return this.percentColumnLetter;
    }

    public int getSheetIndex() {
        return this.sheetIndex;
    }
}