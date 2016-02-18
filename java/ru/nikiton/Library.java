package ru.nikiton;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import java.util.ArrayList;

public class Library {
    public static Cell newCell(Row row, int cellIndex, Double cellValue, boolean fillZero, CellStyle style){
        Cell newCell;
        newCell = row.createCell(cellIndex);
        if(cellValue > 0)
            newCell.setCellValue(cellValue);
        else
            if(fillZero)
                newCell.setCellValue(cellValue);
        newCell.setCellStyle(style);
        return newCell;
    }

    public static Cell newCell(Row row, int cellIndex, Double cellValue, CellStyle style){
        return newCell(row, cellIndex, cellValue, false, style);
    }

    public static Cell newCell(Row row, int cellIndex, String cellValue, CellStyle style){
        Cell newCell;
        newCell = row.createCell(cellIndex);
        newCell.setCellValue(cellValue);
        newCell.setCellStyle(style);
        return newCell;
    }

    public static Cell newCellFormula(Row row, int cellIndex, String formula, CellStyle style){
        Cell newCell;
        newCell = row.createCell(cellIndex);
        newCell.setCellType(Cell.CELL_TYPE_NUMERIC);
        if(!formula.equals(""))
            newCell.setCellFormula(formula);
        newCell.setCellStyle(style);
        return newCell;
    }

    public static Cell newCellFormula(Row row, int cellIndex, String formula){
        Cell newCell;
        CellStyle style;
        if((newCell = row.getCell(cellIndex)) != null) {
            style = newCell.getCellStyle();
            return newCellFormula(row, cellIndex, formula, style);
        }
        else {
            newCell = row.createCell(cellIndex);
            newCell.setCellFormula(formula);
            return newCell;
        }
    }

    public static String getColumnLetter(int columnIndex){
        char[] letters = {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'};
        StringBuilder columnLetterSB = new StringBuilder();
        if(columnIndex == -1 || columnIndex > 255) return columnLetterSB.toString();
        else{
            if(columnIndex > 25){
                Integer index = columnIndex-26;
                columnLetterSB = columnLetterSB.append(letters[(index - index % 26) / 26]);
                columnLetterSB = columnLetterSB.append(letters[index % 26]);
            }
            else
                columnLetterSB = columnLetterSB.append(letters[columnIndex % 26]);
        }
        return columnLetterSB.toString();
    }

    public static String numberToString(Double number, boolean firstLetterIsUp){
        ArrayList<String[]> endings = new ArrayList<String[]>();
        endings.add(new String[] {"рублей", "рубль", "рубля", "рубля", "рубля", "рублей", "рублей", "рублей", "рублей", "рублей"});
        endings.add(new String[] {"тысяч", "тысяча", "тысячи", "тысячи", "тысячи", "тысяч", "тысяч", "тысяч", "тысяч", "тысяч"});

        String[] unitsM = {"error unitsM!", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"};
        String[] unitsG = {"error unitsG!", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"};
        String[] units1 = {"десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"};
        String[] tens = {"error tens!", "десять", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"};
        String[] hundreds = {"error hundreds!", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"};

        String[] summa = number.toString().split("\\.");
        int summaRub = Integer.parseInt(summa[0]);
        if(summa[1].length() == 1)
            summa[1] = summa[1].concat("0");
        ArrayList<Integer> trinities = new ArrayList<Integer>();
        while(summaRub > 0){
            trinities.add(summaRub%1000);
            summaRub = summaRub/1000;
        }
        StringBuilder summByWords = new StringBuilder();
        for(int i=trinities.size()-1; i >= 0; i--){
            Integer nextTrinity = trinities.get(i);
            if(nextTrinity != 0){
                if(i < trinities.size()-1)
                    summByWords.append(" ");
                int last2Digits = nextTrinity%100;
                int firstDigit = nextTrinity/100;
                int secondDigit = last2Digits/10;
                int thirdDigit = last2Digits%10;
                if (firstDigit > 0) {
                    summByWords.append(hundreds[firstDigit]);
                    if (last2Digits != 0)
                        summByWords.append(" ");
                    else {
                        summByWords.append(" ").append(endings.get(i)[0]);
                        continue;
                    }
                }
                if (last2Digits >= 10 && last2Digits < 20){
                    summByWords.append(units1[thirdDigit]);
                    summByWords.append(" ").append(endings.get(i)[0]);
                } else {
                    if (secondDigit != 0){
                        summByWords.append(tens[secondDigit]);
                        if (thirdDigit != 0)
                            summByWords.append(" ");
                    }
                    if (thirdDigit != 0)
                        if (i == 1)//thousands is a female
                            summByWords.append(unitsG[thirdDigit]);
                        else
                            summByWords.append(unitsM[thirdDigit]);
                    summByWords.append(" ").append(endings.get(i)[thirdDigit]);
                }
            }
            else
                summByWords.append(" ").append(endings.get(i)[0]);
        }
        if(!summByWords.toString().equals("")){
            if(firstLetterIsUp){
                Character firstLetter = summByWords.charAt(0);
                summByWords.setCharAt(0, firstLetter.toString().toUpperCase().charAt(0));
            }
            summByWords.append(", ").append(summa[1]).append(" коп.");
        }
        else
            if(Integer.parseInt(summa[1]) > 0){
                summByWords.append(summa[1]).append(" коп.");
            }
            else
                summByWords.append("0 рублей, 0 коп.");

        return summByWords.toString();
    }
}
