package ru.nikiton;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class OperatorWorker {

    public static ArrayList<String> warnings = new ArrayList<String>();

    public static void printFormulaInRow(ArrayList<Operator> operators, String formulaTemplate, Row targetRow, CellStyle style) {
        for(Operator operator : operators) {
            int CI = operator.getColumnIndex();
            String CL = operator.getColumnLetter();
            String sumByPartsFormula = formulaTemplate.replaceAll("x", CL);
            if(style == null)
                Library.newCellFormula(targetRow, CI, sumByPartsFormula);
            else
                Library.newCellFormula(targetRow, CI, sumByPartsFormula, style);
        }
    }

    public static ArrayList<String> make(File file) throws Exception {
        ArrayList<Operator> operators = new ArrayList<Operator>();
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
        wb.setForceFormulaRecalculation(true);
        OperatorCellStyles styles = new OperatorCellStyles(wb);

        //1. Sheet indexes variables
        int mainSheetIndx = -1;
        //2. table indexes variables
        int tableHatRowIndex = -1;
        int tableLastRowIndex = -1;
        int mainSheetLastRowIndex = -1;
        int namesCellIndex = -1;
//        int itogoHoursColumnIndex = -1;
//        int itogoPercentsColumnIndex = -1;

//1. 	Parsing .xls file to sheets + initializing operators Array:
        for(int i=0; i < wb.getNumberOfSheets(); i++) {
            if(wb.getSheetAt(i).getSheetName().equals("Операторы") || wb.getSheetAt(i).getSheetName().equals("Универсалы")) {
                mainSheetIndx = i;
            }
            else {
                if(!wb.getSheetAt(i).getSheetName().equals("Заказы")) {
                    operators.add(new Operator(wb.getSheetAt(i).getSheetName(), i));
                }
            }
        }

//2.    Finding of starting point in Main sheet (index of column with cell "Номер  заказа" and index of row with operators names)
        Sheet mainSheet = wb.getSheetAt(mainSheetIndx);
        Iterator<Row> mainSheetRowIterator = mainSheet.iterator();
        Row currentRowInMain;

        while(mainSheetRowIterator.hasNext()) {
            currentRowInMain = mainSheetRowIterator.next();
            if(tableHatRowIndex == -1) {
                for(int i = currentRowInMain.getFirstCellNum(); i < currentRowInMain.getLastCellNum(); i++) {
                    Cell nextCell = currentRowInMain.getCell(i);
                    if(nextCell != null && nextCell.toString().contains("Номер") && nextCell.toString().contains("заказа")) {
                        namesCellIndex = i;
                        tableHatRowIndex = currentRowInMain.getRowNum();
                    }
/*                    if(nextCell != null && nextCell.toString().contains("ИТОГО") && nextCell.toString().contains("часов")) {
                        itogoHoursColumnIndex = i;
                    }
                    if(nextCell != null && nextCell.toString().contains("ИТОГО") && nextCell.toString().contains("%")) {
                        itogoPercentsColumnIndex = i;
                    }*/
                }
            }
            else {
                Cell nextCell = currentRowInMain.getCell(namesCellIndex);
                if(nextCell != null && nextCell.toString().contains("ВСЕГО") && nextCell.toString().contains("по") && nextCell.toString().contains("заказам")) {
                    tableLastRowIndex = currentRowInMain.getRowNum();
                }
                if(nextCell != null && nextCell.toString().contains("Нормативное") && nextCell.toString().contains("количество") && nextCell.toString().contains("часов")) {
                    mainSheetLastRowIndex = currentRowInMain.getRowNum();
                }
            }
        }
        if(namesCellIndex == -1 || tableHatRowIndex == -1) {
            throw new ExcelWorkerException("В сводной таблице не найдена ячейка \"Номер заказа\"");
        }
/*        if(itogoHoursColumnIndex == -1) {
            throw new ExcelWorkerException("В сводной таблице не найдена ячейка \"ИТОГО часов\"");
        }
        if(itogoPercentsColumnIndex == -1) {
            throw new ExcelWorkerException("В сводной таблице не найдена ячейка \"ИТОГО %\"");
        }*/
        if(tableLastRowIndex == -1) {
            throw new ExcelWorkerException("В сводной таблице не найдена строка \"ВСЕГО по заказам\"");
        }
        if(mainSheetLastRowIndex == -1) {
            throw new ExcelWorkerException("В сводной таблице не найдена строка \"Нормативное количество часов в ... составляет ...\"");
        }

        //2.1. If some rows are already exists in table - remove them
        int lastRowIndex = mainSheet.getLastRowNum();
        int redundantRows = tableHatRowIndex+2-tableLastRowIndex; //negative number

        if(redundantRows < 0) {
            //Remove all rows in table
            for(int i = tableHatRowIndex+2; i < tableLastRowIndex; i++) {
                mainSheet.removeRow(mainSheet.getRow(i));
            }
            //Collapsing table
            mainSheet.shiftRows(tableLastRowIndex, lastRowIndex, redundantRows, true, true);
        }
        //2.2. Remove redundant rows under the table if they exists
        for(int i = lastRowIndex; i > mainSheetLastRowIndex; i--)
            if(mainSheet.getRow(i) != null)
                mainSheet.removeRow(mainSheet.getRow(i));

//3.	Updating operators workingTimeMap data:
        //3.1. filling operators workingTimeMaps by real data and create fullOrdersMap - TreeMap of all orders, order parts and its mtp in operators sheets
        TreeMap<String, TreeMap<String, TreeSet<String>>> fullOrdersMap = new TreeMap<String, TreeMap<String, TreeSet<String>>>();
        for(Operator operator : operators) {
            Sheet operatorSheet = wb.getSheetAt(operator.getSheetIndex());
            boolean endData = false;
            int operatorSheetRowIndex = 2;
            //working with each operator sheets, calculating operators worktime and saving data in workingTimeMap
            while(!endData) {
                Row nextRow = operatorSheet.getRow(operatorSheetRowIndex);
                String orderNum = nextRow.getCell(6).toString();
                String partName = nextRow.getCell(7).toString();
                //Cast MTP cell to String type and get right data
                Cell mtpCell = nextRow.getCell(8);
                mtpCell.setCellType(Cell.CELL_TYPE_STRING);
                String mtp = mtpCell.getStringCellValue();
                if(orderNum == null || "".equals(orderNum) || partName == null) {
                    endData = true;
                }
                else {
                    if(!orderNum.equals("200.100") && (mtp == null || mtp.equals(""))) {
                        warnings.add("У работника " + operator.getName() + " не заполнено МТП в строке №" + (operatorSheetRowIndex+1));
                    }
                    if(operator.updateWorkingTimeMap(orderNum, partName, nextRow.getCell(5).getNumericCellValue(), mtp)) {
                        TreeMap<String, TreeSet<String>> currentOrderContent = fullOrdersMap.get(orderNum);
                        if(currentOrderContent != null) {
                            TreeSet<String> currentPartMTPs = currentOrderContent.get(partName);
                            if(currentPartMTPs != null) {
                                currentPartMTPs.add(mtp);
                            }
                            else {
                                TreeSet<String> newMTP = new TreeSet<String>();
                                newMTP.add(mtp);
                                currentOrderContent.put(partName, newMTP);
                            }
                        }
                        else {
                            TreeSet<String> newMTP = new TreeSet<String>();
                            newMTP.add(mtp);
                            TreeMap<String, TreeSet<String>> newOrderContent = new TreeMap<String, TreeSet<String>>();
                            newOrderContent.put(partName, newMTP);
                            fullOrdersMap.put(orderNum, newOrderContent);
                        }
                    }
                    operatorSheetRowIndex++;
                }
            }
            //searching operators names in Main sheet and saving it`s cell positions in operator`s mainPos property
            Row tableHatRow = mainSheet.getRow(tableHatRowIndex);
            //for(int i = namesCellIndex+1; i < itogoHoursColumnIndex; i++) {
            for(int i = namesCellIndex+1; i < tableHatRow.getLastCellNum(); i++) {
                if(tableHatRow.getCell(i).toString().contains(operator.getName())) {
                    operator.setColumnIndex(i, Library.getColumnLetter(i));
                    break;
                }
            }
        }

        //3.2. calculating number of rows in table
        int countTableRows = fullOrdersMap.size();
        for(Map.Entry<String, TreeMap<String, TreeSet<String>>> fwtOrderData : fullOrdersMap.entrySet()) {
            if(!fwtOrderData.getKey().equals("200.100")) {
                TreeMap<String, TreeSet<String>> fwtPart = fwtOrderData.getValue();
                countTableRows += fwtPart.size();
                for(Map.Entry<String, TreeSet<String>> fwtPartData : fwtPart.entrySet()) {
                    countTableRows += fwtPartData.getValue().size();
                }
            }
        }

        if(countTableRows == 0) {
            wb.write(new FileOutputStream(file));
            wb.close();
            throw new ExcelWorkerException("Отсутствуют данные на страницах операторов!");
        }

        //3.3. expand operators workingtime maps to fullOrdersMap and
        for(Operator operator : operators) {
            for(Map.Entry<String, TreeMap<String, TreeSet<String>>> fwtOrderData : fullOrdersMap.entrySet()) {
                for(Map.Entry<String, TreeSet<String>> fwtPartData : fwtOrderData.getValue().entrySet()) {
                    for(String mtp : fwtPartData.getValue())
                        operator.updateWorkingTimeMap(fwtOrderData.getKey(), fwtPartData.getKey(), 0d, mtp);
                }
            }
        }

//4.    Updating Main sheet
        //4.1. Add necessary rows in table
        int tableFirstRowIndex = tableHatRowIndex+2;
        //Update tableLastRowIndex after row calculating
        tableLastRowIndex = tableFirstRowIndex+countTableRows;
        //Create additional rows for correct shifting
        for(int i=countTableRows-(mainSheet.getLastRowNum()+1-tableFirstRowIndex); i>0; i--) {
            mainSheet.createRow(mainSheet.getLastRowNum()+1);
        }
        //Shift lower table rows to creating new rows in the middle of table
        mainSheet.shiftRows(tableFirstRowIndex, mainSheet.getLastRowNum(), countTableRows, true, true);

        //4.2. Print all orders and parts names in 1st column of table + print all formulas in table
        int currentRowIndex = tableFirstRowIndex;
        String sumByOrdersFormulaTemplate = "";
        for(Map.Entry<String, TreeMap<String, TreeSet<String>>> order : fullOrdersMap.entrySet()) {
            String orderName = order.getKey();
            TreeMap<String, TreeSet<String>> orderContent = order.getValue();
            //print order name, its parts names and MTPs in 1st column of table
            Row orderNextRow = mainSheet.getRow(currentRowIndex);
            orderNextRow.setHeightInPoints(20);
            Cell newCell = orderNextRow.createCell(namesCellIndex);
            newCell.setCellValue(orderName);
            newCell.setCellStyle(styles.orderNameStyle);
            //Create cells in last 2 columns (for correct styling)
/*            Cell lastColValueCell = orderNextRow.createCell(itogoHoursColumnIndex);
            lastColValueCell.setCellStyle(styles.orderValStyle);
            Cell lastColPrcCell = orderNextRow.createCell(itogoPercentsColumnIndex);
            lastColPrcCell.setCellStyle(styles.orderPrcStyle);
*/
            currentRowIndex++;
            sumByOrdersFormulaTemplate += "x"+(currentRowIndex)+"+";
            if(!orderName.equals("200.100")) {
                String sumByPartsFormulaTemplate = "";
                for(Map.Entry<String, TreeSet<String>> partData : orderContent.entrySet()) {
                    Row partNextRow = mainSheet.getRow(currentRowIndex);
                    partNextRow.setHeightInPoints(20);
                    newCell = partNextRow.createCell(namesCellIndex);
                    newCell.setCellValue(partData.getKey());
                    newCell.setCellStyle(styles.partNameStyle);
                    //Create cells in last 2 columns (for correct styling)
/*                    lastColValueCell = partNextRow.createCell(itogoHoursColumnIndex);
                    lastColValueCell.setCellStyle(styles.partValStyle);
                    lastColPrcCell = partNextRow.createCell(itogoPercentsColumnIndex);
                    lastColPrcCell.setCellStyle(styles.partPrcStyle);
*/
                    currentRowIndex++;
                    sumByPartsFormulaTemplate += "x"+(currentRowIndex)+"+";
                    String sumByMTPsFormulaTemplate = "";
                    for(String mtp : partData.getValue()) {
                        Row mtpNextRow = mainSheet.getRow(currentRowIndex);
                        mtpNextRow.setHeightInPoints(20);
                        newCell = mtpNextRow.createCell(namesCellIndex);
                        if(!mtp.equals("")) {
                            newCell.setCellValue(mtp);
                            newCell.setCellStyle(styles.mtpNameStyle);
                        }
                        else {
                            newCell.setCellValue("Нет МТП");
                            newCell.setCellStyle(styles.mtpNameErrorStyle);
                        }
                        //Create cells in last 2 columns (for correct styling)
/*                        lastColValueCell = mtpNextRow.createCell(itogoHoursColumnIndex);
                        lastColValueCell.setCellStyle(styles.partValStyle);
                        lastColPrcCell = mtpNextRow.createCell(itogoPercentsColumnIndex);
                        lastColPrcCell.setCellStyle(styles.partPrcStyle);
*/
                        currentRowIndex++;
                        sumByMTPsFormulaTemplate += "x"+(currentRowIndex)+"+";
                    }
                    sumByMTPsFormulaTemplate = sumByMTPsFormulaTemplate.substring(0, sumByMTPsFormulaTemplate.length()-1);
                    printFormulaInRow(operators, sumByMTPsFormulaTemplate, partNextRow, styles.partValStyle);
                }
                sumByPartsFormulaTemplate = sumByPartsFormulaTemplate.substring(0, sumByPartsFormulaTemplate.length()-1);
                printFormulaInRow(operators, sumByPartsFormulaTemplate, orderNextRow, styles.orderValStyle);
            }
        }
        //Make formula template for row with total amount formulas ("ВСЕГО по заказам")
        sumByOrdersFormulaTemplate = sumByOrdersFormulaTemplate.substring(0, sumByOrdersFormulaTemplate.length()-1);
        // Declare formulas for columns "ИТОГО", "ИТОГО %"
//        String summAllOperatorsFormulaTemplate = "";

        //4.3. Fill table columns "к-во времени, час" by operators
        for(Operator operator : operators) {
            TreeMap<String, TreeMap<String, TreeMap<String, Double>>> workingTimeMap = operator.getWorkingTimeMap();
            int CI = operator.getColumnIndex();
            currentRowIndex = tableFirstRowIndex;
            for(Map.Entry<String, TreeMap<String, TreeMap<String, Double>>> orderData : workingTimeMap.entrySet()) {
                Row orderRow = mainSheet.getRow(currentRowIndex);
                CellReference orderHoursCellRef = new CellReference(currentRowIndex, CI);
                CellReference totalOrderHoursCellRef = new CellReference(tableLastRowIndex, CI);
                //String worktimeInPrcFormula = "ROUND((".concat(orderHoursCellRef.formatAsString()).concat("*100)/").concat(totalOrderHoursCellRef.formatAsString()).concat(",0)");
                String worktimeInPrcFormula = "ROUND((".concat(orderHoursCellRef.formatAsString()).concat("*100)/").concat(totalOrderHoursCellRef.formatAsString()).concat(",2)");
                if(orderData.getKey().equals("200.100")) {
                    Library.newCell(orderRow, CI, orderData.getValue().get("").get(""), styles.orderValStyle);
                    //Set percents formula in next cell
                    Library.newCellFormula(orderRow, CI+1, worktimeInPrcFormula, styles.orderPrcStyle);
                    currentRowIndex++;
                    continue;
                }
                //Set order workingtime in percents formula
                Library.newCellFormula(orderRow, CI+1, worktimeInPrcFormula, styles.orderPrcStyle);
                currentRowIndex++;
                //Fill worktime hours in mtp cells:
                for(Map.Entry<String, TreeMap<String, Double>> partData : orderData.getValue().entrySet()) {
                    //Print parts data (percents only)
                    //Create and set worktime in percents
                    CellReference partHoursCellRef = new CellReference(currentRowIndex, CI);
                    CellReference totalPartHoursCellRef = new CellReference(tableLastRowIndex, CI);
                    //worktimeInPrcFormula = "ROUND((".concat(partHoursCellRef.formatAsString()).concat("*100)/").concat(totalPartHoursCellRef.formatAsString()).concat(",0)");
                    worktimeInPrcFormula = "ROUND((".concat(partHoursCellRef.formatAsString()).concat("*100)/").concat(totalPartHoursCellRef.formatAsString()).concat(",2)");
                    Row partRow = mainSheet.getRow(currentRowIndex);
                    Library.newCellFormula(partRow, CI+1, worktimeInPrcFormula, styles.partPrcStyle);
                    currentRowIndex++;
                    for(Map.Entry<String, Double> mtpData : partData.getValue().entrySet()) {
                        Row mtpRow = mainSheet.getRow(currentRowIndex);
                        //Print worktime data and set worktime in percents
                        Double mtpVal = mtpData.getValue();
                        //Create worktime in percents formula
                        CellReference mtpHoursCellRef = new CellReference(currentRowIndex, CI);
                        CellReference totalMtpHoursCellRef = new CellReference(tableLastRowIndex, CI);
                        //worktimeInPrcFormula = "ROUND((".concat(mtpHoursCellRef.formatAsString()).concat("*100)/").concat(totalMtpHoursCellRef.formatAsString()).concat(",0)");
                        worktimeInPrcFormula = "ROUND((".concat(mtpHoursCellRef.formatAsString()).concat("*100)/").concat(totalMtpHoursCellRef.formatAsString()).concat(",2)");

                        Library.newCell(mtpRow, CI, mtpVal, true, styles.mtpValStyle);
                        Library.newCellFormula(mtpRow, CI+1, worktimeInPrcFormula, styles.mtpPrcStyle);

                        currentRowIndex++;
                    }
                }
            }
            //Set 2 formulas in each operators columns "ВСЕГО по заказам"
            int CI1 = operator.getPercentColumnIndex();
            String CL = operator.getColumnLetter();
            String CL1 = operator.getPercentColumnLetter();
            String sumByPartsFormula = sumByOrdersFormulaTemplate.replaceAll("x", CL);
            Library.newCellFormula(mainSheet.getRow(tableLastRowIndex), CI, sumByPartsFormula);
            String sumByPartsInPrcFormula = sumByOrdersFormulaTemplate.replaceAll("x", CL1);
            Library.newCellFormula(mainSheet.getRow(tableLastRowIndex), CI1, sumByPartsInPrcFormula);
            //And assemble formula templates for last 2 columns (ИТОГО, ИТОГО %)
//            summAllOperatorsFormulaTemplate += CL+"x"+"+";
        }
        //Make formula template for column ("ИТОГО часов")
/*        summAllOperatorsFormulaTemplate = summAllOperatorsFormulaTemplate.substring(0, summAllOperatorsFormulaTemplate.length()-1);

        //4.4. Set formulas in last 2 columns "ИТОГО часов", "ИТОГО %"
        currentRowIndex = tableFirstRowIndex;
        while(currentRowIndex < tableLastRowIndex) {
            Row currentRow = mainSheet.getRow(currentRowIndex);
            String RN = ((Integer)(currentRowIndex+1)).toString(); //RN - Row Number in real excel indexing (1 to 65535)
            String summAllOperatorsFormula = summAllOperatorsFormulaTemplate.replaceAll("x", RN);
            //"ИТОГО часов" formula
            //Cell itogoHoursCell = currentRow.createCell(itogoHoursColumnIndex);
            Cell itogoHoursCell = currentRow.getCell(itogoHoursColumnIndex);
            itogoHoursCell.setCellFormula(summAllOperatorsFormula);
            //"ИТОГО %" formula
            String itogoCL = Library.getColumnLetter(itogoHoursColumnIndex);
            String summAllOperatorsInPrcFormula = "ROUND((".concat(itogoCL).concat(RN).concat("*100)/").concat(itogoCL).concat("" + (tableLastRowIndex+1)).concat(",0)");
            //Cell itogoPercentsCell = currentRow.createCell(itogoPercentsColumnIndex);
            Cell itogoPercentsCell = currentRow.getCell(itogoPercentsColumnIndex);
            //itogoPercentsCell.setCellStyle(currentRow.getCell(itogoPercentsColumnIndex-2).getCellStyle());
            itogoPercentsCell.setCellFormula(summAllOperatorsInPrcFormula);
            currentRowIndex++;
        }
        //Set formula in "ВСЕГО по заказам" by % FINAL cell
        String finalSummByPercentsFormula = sumByOrdersFormulaTemplate.replaceAll("x", Library.getColumnLetter(itogoPercentsColumnIndex));
        Row currentRow = mainSheet.getRow(currentRowIndex);
        Cell itogoPercentsCell = currentRow.getCell(itogoPercentsColumnIndex);
        itogoPercentsCell.setCellFormula(finalSummByPercentsFormula);
*/
//5. Saving changes into file
        wb.write(new FileOutputStream(file));
        wb.close();
        return warnings;
    }
}