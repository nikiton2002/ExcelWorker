package ru.nikiton;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.*;
import java.util.*;

public class ContractorWorker {
    public static void make(File file) throws Exception {
        ArrayList<Contractor> contractors = new ArrayList<Contractor>();
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
        wb.setForceFormulaRecalculation(true);  //Very IMPORTANT COMMAND! It says to Excel to refresh all formulas when file opens
        ContractorCellStyles styles = new ContractorCellStyles(wb);

//1. 	Parsing .xls file to sheets + initializing operators Array:
        int countWbSheets = wb.getNumberOfSheets();
        for(int i=0; i < countWbSheets; i++) {
            String contractorName;
            int mainSheetIndex;
            int worktimeSheetIndex = -1;
            String sheetName = wb.getSheetAt(i).getSheetName();
            if(sheetName.contains("Акт") && sheetName.contains("-")) {
                contractorName = sheetName.substring(0, sheetName.indexOf("-")).trim();
                mainSheetIndex = i;
                for(int j=0; j < countWbSheets; j++) {
                    if(wb.getSheetAt(j).getSheetName().contains(contractorName) &&
                            (wb.getSheetAt(j).getSheetName().contains("табель") || wb.getSheetAt(j).getSheetName().contains("работа"))) {
                        worktimeSheetIndex = j;
                        break;
                    }
                }
                if(worktimeSheetIndex == -1) {
                    throw new ExcelWorkerException("В файле найден Акт подрядчика "+contractorName+", но не найдена его страница рабочего времени!");
                }
                contractors.add(new Contractor(contractorName, worktimeSheetIndex, mainSheetIndex));
            }
        }

//2.    Updating contractors data:
        for(Contractor contractor : contractors) {
            //2.1. Find and save all column indexes and index of hat row of contractor`s table in main sheet
            int tableHatRowIndex = -1;  //index of hat row of contractor`s table
            int tableFootRowIndex = -1;  //index of foot row of contractor`s table (Итого)
            int punktNumCellIndex = -1;  //index of column with cell "№ п/п"
            int serviceNameCellIndex = -1;  //index of column with cell "Наименование работ"
            int orderCellIndex = -1;  //index of column with cell "№ заказа"
            int workHoursCellIndex = -1;  //index of column with cell "Количество отработанных часов"
            int serviceTariffCellIndex = -1;  //index of column with cell "Тариф с НДФЛ, руб."
            int serviceCostCellIndex = -1;  //index of column with cell "Стоимость работ с НДФЛ, руб."

            HSSFSheet mainSheet = wb.getSheetAt(contractor.getMainSheetIndex());
            Iterator<Row> mainSheetRowIterator = mainSheet.iterator();
            while(mainSheetRowIterator.hasNext()) {
                Row currentRow = mainSheetRowIterator.next();
                if(tableHatRowIndex == -1) {
                    for(int i = currentRow.getFirstCellNum(); i < currentRow.getLastCellNum(); i++) {
                        Cell nextCell = currentRow.getCell(i);
                        if(nextCell != null && nextCell.toString().contains("№") && nextCell.toString().contains("п/п")) {
                            punktNumCellIndex = i;
                        }
                        if(nextCell != null && nextCell.toString().contains("Наименование") && nextCell.toString().contains("работ")) {
                            serviceNameCellIndex = i;
                        }
                        if(nextCell != null && nextCell.toString().contains("№") && nextCell.toString().contains("заказа")) {
                            orderCellIndex = i;
                            tableHatRowIndex = currentRow.getRowNum();
                        }
                        if(nextCell != null && nextCell.toString().contains("Количество") && nextCell.toString().contains("отработанных") && nextCell.toString().contains("часов")) {
                            workHoursCellIndex = i;
                        }
                        if(nextCell != null && nextCell.toString().contains("Тариф") && nextCell.toString().contains("НДФЛ") && nextCell.toString().contains("руб")) {
                            serviceTariffCellIndex = i;
                        }
                        if(nextCell != null && nextCell.toString().contains("Стоимость") && nextCell.toString().contains("НДФЛ") && nextCell.toString().contains("руб")) {
                            serviceCostCellIndex = i;
                        }
                    }
                }
                else{
                    Cell nextCell = currentRow.getCell(punktNumCellIndex);
                    if(nextCell != null && nextCell.toString().contains("Итого") && nextCell.toString().contains(":")) {
                        tableFootRowIndex = currentRow.getRowNum();
                    }
                }
                if(tableHatRowIndex > -1 && tableFootRowIndex > -1) {
                    break;
                }
            }
            if(punktNumCellIndex == -1 || serviceNameCellIndex == -1 || orderCellIndex == -1 || workHoursCellIndex == -1 || serviceTariffCellIndex == -1 || serviceCostCellIndex == -1 || tableHatRowIndex == -1) {
                throw new ExcelWorkerException("В Акте подрядчика "+contractor.getName()+" не найдена строка с шапкой таблицы! В шапке таблицы должны присутствовать ячейки: \"№ п/п\", \"Наименование работ\", \"№ заказа\", \"Количество отработанных часов\", \"Тариф с НДФЛ, руб\", \"Стоимость работ с НДФЛ, руб\"");
            }
            if(tableFootRowIndex == -1) {
                throw new ExcelWorkerException("В Акте подрядчика \"+contractor.getName()+\" не найдена строка Итого!");
            }
            //2.2. If some rows already exists in table - delete them
            if(tableFootRowIndex > tableHatRowIndex+2) {
                int tableRowIndex = tableHatRowIndex+2;
                CellRangeAddress mergedRegion;
                //Remove merged regions in table
                while(tableRowIndex < tableFootRowIndex) {
                    Row currentRow = mainSheet.getRow(tableRowIndex);
                    for(int i = currentRow.getFirstCellNum(); i < currentRow.getLastCellNum(); i++) {
                        Cell nextCell = currentRow.getCell(i);
                        if(nextCell == null)
                            continue;
                        int nextCellCI = nextCell.getColumnIndex();
                        int nextCellRI = nextCell.getRowIndex();
                        for(int j=0; j < mainSheet.getNumMergedRegions(); j++) {
                            mergedRegion = mainSheet.getMergedRegion(j);
                            if(mergedRegion.getFirstColumn() <= nextCellCI && mergedRegion.getLastColumn() >= nextCellCI
                                    && mergedRegion.getFirstRow() <= nextCellRI && mergedRegion.getLastRow() >= nextCellRI) {
                                mainSheet.removeMergedRegion(j);
                            }
                        }
                    }

                    tableRowIndex++;
                }
                //Remove all rows in table
                for(int i = tableHatRowIndex+2; i < tableFootRowIndex; i++) {
                    mainSheet.removeRow(mainSheet.getRow(i));
                }
                //Collapsing table
                int collapsedRows = tableFootRowIndex-(tableHatRowIndex+2);
                mainSheet.shiftRows(tableFootRowIndex, mainSheet.getLastRowNum(), -collapsedRows, false, true);
                //Update tableFootRowIndex
                tableFootRowIndex -= collapsedRows;
            }

            //2.3. Fill contractor`s worktime map
            Sheet worktimeSheet = wb.getSheetAt(contractor.getSheetIndex());
            //Saving contractor`s worktime price and services
            Row currentRow = worktimeSheet.getRow(1);
            boolean endData = false;
            boolean singleService = false;
            int worktimeTableRowIndex = 2;
            String serviceName = "";
            if(!"".equals(currentRow.getCell(8).toString())) {
                serviceName = currentRow.getCell(8).toString();
                if(currentRow.getCell(9) != null &&
                   currentRow.getCell(9).getCellType() == Cell.CELL_TYPE_NUMERIC &&
                   currentRow.getCell(9).getNumericCellValue() > 0)
                {
                    contractor.setHourPrice(serviceName, currentRow.getCell(9).getNumericCellValue());
                }
                else
                    contractor.setHourPrice(serviceName, 0d);
                singleService = true;
            }
            //Reading contractor`s worktime-sheet and saving data in workingTimeMap
            while(!endData) {
                currentRow = worktimeSheet.getRow(worktimeTableRowIndex);
                Double workingTime = currentRow.getCell(5).getNumericCellValue();
                String orderNum = currentRow.getCell(6).toString();
                String partName = currentRow.getCell(7).toString();
                //Cast MTP cell to String type and get right data
                Cell mtpCell = currentRow.getCell(10);
                mtpCell.setCellType(Cell.CELL_TYPE_STRING);
                String mtp = mtpCell.getStringCellValue();

                if(orderNum == null || "".equals(orderNum) || partName == null)
                {
                    endData = true;
                }
                else {
                    if(!singleService) {
                        if(currentRow.getCell(8) != null) {
                            String currentServiceName = currentRow.getCell(8).getStringCellValue();
                            if(!"".equals(currentServiceName))
                                serviceName = currentServiceName;
                        }
                        if(currentRow.getCell(9) != null &&
                           currentRow.getCell(9).getCellType() == Cell.CELL_TYPE_NUMERIC &&
                           currentRow.getCell(9).getNumericCellValue() > 0)
                        {
                            contractor.setHourPrice(serviceName, currentRow.getCell(9).getNumericCellValue());
                        }
                        else {
                            if(contractor.getHourPrice(serviceName) == null)
                                contractor.setHourPrice(serviceName, 0d);
                        }
                    }
                    contractor.updateWorkingTimeMap(serviceName, orderNum, partName, mtp, workingTime);
                    worktimeTableRowIndex++;
                }
            }

//3.    Updating contractor`s Main sheets
            TreeMap<String, TreeMap<String, TreeMap<String, TreeMap<String, Double>>>> contractorWorkingTimeMap = contractor.getWorkingTimeMap();
            //No data -> no map, no map -> no working -> going to next contractor
            if(contractorWorkingTimeMap.isEmpty())
                continue;
            int serviceNumInTable=1;
            //Print orders by services
            int countNewRows;
            int serviceLastRowIndex = tableHatRowIndex+2; //initialization tableLastRowIndex for correct adding first service in table
            //Initialize formula: sum by ALL contractor`s orders
            String contractorsWorktimeSumFormula = "";
            String contractorsTotalCostSumFormula = "";
            for(Map.Entry<String, TreeMap<String, TreeMap<String, TreeMap<String, Double>>>> serviceData : contractorWorkingTimeMap.entrySet()) {
                serviceName = serviceData.getKey();
                TreeMap<String, TreeMap<String, TreeMap<String, Double>>> serviceWorkingTimeMap = serviceData.getValue();
                int serviceFirstRowIndex = serviceLastRowIndex; //serviceFirstRowIndex is an index of first row in table for current service

                // 3.1. calculating number of new rows needed in working time table for current service
                countNewRows = serviceWorkingTimeMap.size(); //number of orders in service
                for(Map.Entry<String, TreeMap<String, TreeMap<String, Double>>> orderData : serviceWorkingTimeMap.entrySet()) {
                    if(!orderData.getKey().equals("200.100")) {
                        countNewRows += orderData.getValue().size();
                        for(Map.Entry<String, TreeMap<String, Double>> partData : orderData.getValue().entrySet()) {
                            countNewRows += partData.getValue().size();
                        }
                    }
                }

                //3.2. Add necessary rows in table for current service
                serviceLastRowIndex = serviceFirstRowIndex + countNewRows;
                //Create additional rows for correct shifting
                for(int i=countNewRows - (mainSheet.getLastRowNum()+1 - serviceFirstRowIndex); i>0; i--) {
                    mainSheet.createRow(mainSheet.getLastRowNum()+1);
                }
                //Shift lower table rows to creating new rows in the middle of table
                mainSheet.shiftRows(serviceFirstRowIndex, mainSheet.getLastRowNum(), countNewRows, false, true);
                //Update tableFootRowIndex
                tableFootRowIndex += countNewRows;
                //3.3. Printing data of current service in table
                //Set static values "№ п/п", "Наименование работ" and "Тариф с НДФЛ, руб." (1st, 2nd and 5th columns)
                Row firstRow = mainSheet.getRow(serviceFirstRowIndex);
                Library.newCell(firstRow, punktNumCellIndex, serviceNumInTable + ".", styles.orderValStyle); //Number of service in table
                serviceNumInTable++;
                Library.newCell(firstRow, serviceNameCellIndex, serviceName, styles.orderValStyle);         //Name of service
                Library.newCell(firstRow, serviceTariffCellIndex, contractor.getHourPrice(serviceName), true, styles.itogoValStyle);         //Service`s price
                //Make region for "№ п/п" column
                mainSheet.addMergedRegion(new CellRangeAddress(serviceFirstRowIndex, serviceLastRowIndex-1, punktNumCellIndex, punktNumCellIndex));
                //Make region for "Наименование работ" column (it is double-column)
                CellRangeAddress serviceNameCells = new CellRangeAddress(serviceFirstRowIndex, serviceLastRowIndex-1, serviceNameCellIndex, serviceNameCellIndex+1);
                mainSheet.addMergedRegion(serviceNameCells);
                RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM, serviceNameCells, mainSheet, wb);
                RegionUtil.setBorderLeft(CellStyle.BORDER_MEDIUM, serviceNameCells, mainSheet, wb);
                RegionUtil.setBorderRight(CellStyle.BORDER_MEDIUM, serviceNameCells, mainSheet, wb);
                //Make region for "Тариф с НДФЛ, руб." column
                CellRangeAddress servicePriceCells = new CellRangeAddress(serviceFirstRowIndex, serviceLastRowIndex-1, serviceTariffCellIndex, serviceTariffCellIndex);
                mainSheet.addMergedRegion(servicePriceCells);
                RegionUtil.setBorderRight(CellStyle.BORDER_MEDIUM, servicePriceCells, mainSheet, wb);

                //Print orders data
                int currentRowIndex = serviceFirstRowIndex;
                for(Map.Entry<String, TreeMap<String, TreeMap<String, Double>>> orderData : serviceWorkingTimeMap.entrySet()) {
                    //Print order nubmer and its formula
                    Row orderRow = mainSheet.getRow(currentRowIndex);
                    orderRow.setHeightInPoints(20);
                    //Assembling formulas: sum by ALL contractor`s orders  and sum by ALL contractor`s costs
                    CellReference orderHoursCellRef = new CellReference(currentRowIndex, workHoursCellIndex);
                    contractorsWorktimeSumFormula += orderHoursCellRef.formatAsString() + "+";
                    CellReference orderCostCellRef = new CellReference(currentRowIndex, serviceCostCellIndex);
                    contractorsTotalCostSumFormula += orderCostCellRef.formatAsString() + "+";
                    //Print order`s data
                    String orderName = orderData.getKey();
                    Library.newCell(orderRow, orderCellIndex, orderName, styles.orderNameStyle);    //Number of order
                    CellRangeAddress orderNumCells = new CellRangeAddress(currentRowIndex, currentRowIndex, orderCellIndex, orderCellIndex+1);
                    mainSheet.addMergedRegion(orderNumCells);
                    RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM, orderNumCells, mainSheet, wb);
                    RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, orderNumCells, mainSheet, wb);
                    CellReference orderPriceCellRef = new CellReference(serviceFirstRowIndex, serviceTariffCellIndex);
                    if(orderName.equals("200.100")) {
                        //Set hours value
                        Library.newCell(orderRow, workHoursCellIndex, orderData.getValue().get("").get(""), styles.mtpValStyle);
                        //Assembling and set order payment formula in order cost cell (order hours * order hour price)
                        String orderItogFormula = orderHoursCellRef.formatAsString() + "*" + orderPriceCellRef.formatAsString();
                        Library.newCellFormula(orderRow, serviceCostCellIndex, orderItogFormula, styles.itogoValStyle);
                        currentRowIndex++;
                        continue;
                    }
                    //Initialize sum by parts formula
                    String sumByPartsFormula = "";
                //Print parts rows of current order
                    currentRowIndex++;
                    for(Map.Entry<String, TreeMap<String, Double>> partData : orderData.getValue().entrySet()) {
                        Row partRow = mainSheet.getRow(currentRowIndex);
                        partRow.setHeightInPoints(20);
                        Library.newCell(partRow, orderCellIndex, partData.getKey(), styles.partNameStyle);    //Name of order`s part
                        CellRangeAddress partNameCells = new CellRangeAddress(currentRowIndex, currentRowIndex, orderCellIndex, orderCellIndex+1);
                        mainSheet.addMergedRegion(partNameCells);
                        RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, partNameCells, mainSheet, wb);
                        //Assemble sum by parts formula
                        CellReference cellRef = new CellReference(currentRowIndex, workHoursCellIndex);
                        sumByPartsFormula += cellRef.formatAsString() + "+";
                        //Initialize sum by MTP formula
                        String sumByMTPFormula = "";
                        currentRowIndex++;
                    //Print MTPs rows of current part
                        for(Map.Entry<String, Double> mtpData : partData.getValue().entrySet()) {
                            Row mtpRow = mainSheet.getRow(currentRowIndex);
                            mtpRow.setHeightInPoints(20);
                            //Create MTP-name cell
                            Library.newCell(mtpRow, orderCellIndex, mtpData.getKey(), styles.mtpNameStyle);
                            //Makeup for MTP-name cell
                            CellRangeAddress mtpCells = new CellRangeAddress(currentRowIndex, currentRowIndex, orderCellIndex, orderCellIndex+1);
                            mainSheet.addMergedRegion(mtpCells);
                            RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, mtpCells, mainSheet, wb);
                            //Create MTP-value cell
                            Library.newCell(mtpRow, workHoursCellIndex, mtpData.getValue(), styles.mtpValStyle);
                            //Set formula in cell "Стоимость работ с НДФЛ, руб. " by MTP (order hours * order hour price)
                            CellReference mtpHoursCellRef = new CellReference(currentRowIndex, workHoursCellIndex);
                            String mtpItogFormula = mtpHoursCellRef.formatAsString() + "*" + orderPriceCellRef.formatAsString();
                            Library.newCellFormula(mtpRow, serviceCostCellIndex, mtpItogFormula, styles.mtpValStyle);
                            //Assemble sum by MTP formula
                            cellRef = new CellReference(currentRowIndex, workHoursCellIndex);
                            sumByMTPFormula += cellRef.formatAsString() + "+";
                            currentRowIndex++;
                        }
                        //Set sum by MTPs formula
                        sumByMTPFormula = sumByMTPFormula.substring(0, sumByMTPFormula.length()-1);
                        Library.newCellFormula(partRow, workHoursCellIndex, sumByMTPFormula, styles.partValStyle);
                        //Set formula in cell "Стоимость работ с НДФЛ, руб. " by part (order hours * order hour price)
/*                        CellReference partHoursCellRef = new CellReference(partRow.getRowNum(), workHoursCellIndex);
                        String partItogFormula = partHoursCellRef.formatAsString() + "*" + orderPriceCellRef.formatAsString();
                        Library.newCellFormula(partRow, serviceCostCellIndex, partItogFormula, styles.partValStyle);*/
                        Library.newCell(partRow, serviceCostCellIndex, "", styles.partValStyle);
                    }
                    //Set sum by parts formula
                    sumByPartsFormula = sumByPartsFormula.substring(0, sumByPartsFormula.length()-1);
                    Library.newCellFormula(orderRow, workHoursCellIndex, sumByPartsFormula, styles.orderValStyle);
                    //Set formula in order cost cell "Стоимость работ с НДФЛ, руб. " (order hours * order hour price)
                    String orderItogFormula = orderHoursCellRef.formatAsString() + "*" + orderPriceCellRef.formatAsString();
                    Library.newCellFormula(orderRow, serviceCostCellIndex, orderItogFormula, styles.orderValStyle);
                    //Make region for "Стоимость работ с НДФЛ, руб." column of current order
/*                    int orderRowIndex = orderRow.getRowNum();
                    CellRangeAddress orderCostCells = new CellRangeAddress(orderRowIndex, currentRowIndex-1, serviceCostCellIndex, serviceCostCellIndex);
                    mainSheet.addMergedRegion(orderCostCells);
                    RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM, orderCostCells, mainSheet, wb);
                    RegionUtil.setBorderRight(CellStyle.BORDER_MEDIUM, orderCostCells, mainSheet, wb);*/
                }
                currentRow = mainSheet.getRow(currentRowIndex);
            }
            //Set TOTAL formulas:
            contractorsWorktimeSumFormula = contractorsWorktimeSumFormula.substring(0, contractorsWorktimeSumFormula.length() - 1);
            Library.newCellFormula(currentRow, workHoursCellIndex, contractorsWorktimeSumFormula, styles.itogoHrsValStyle);
            contractorsTotalCostSumFormula = contractorsTotalCostSumFormula.substring(0, contractorsTotalCostSumFormula.length()-1);
            Library.newCellFormula(currentRow, serviceCostCellIndex, contractorsTotalCostSumFormula, styles.itogoValStyle);
            currentRow.setHeightInPoints(20);
            //3.4. Find under-table formulas, calculate them and write numbers by words
            //Searching formulas:
            int currentRowIndex = tableFootRowIndex+1;
            while(currentRowIndex <= mainSheet.getLastRowNum()) {
                currentRow = mainSheet.getRow(currentRowIndex);
                for(int i = currentRow.getFirstCellNum(); i < currentRow.getLastCellNum(); i++) {
                    Cell nextCell = currentRow.getCell(i);
                    if(nextCell != null && nextCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                        currentRow.setHeightInPoints(20);
                        FormulaEvaluator calculon = wb.getCreationHelper().createFormulaEvaluator();
                        calculon.evaluateFormulaCell(nextCell);
                        String inWords = Library.numberToString(nextCell.getNumericCellValue(), true);
                        for(int j = i+1; j < currentRow.getLastCellNum(); j++) {
                            nextCell = currentRow.getCell(j);
                            if(nextCell != null && nextCell.getCellType() == Cell.CELL_TYPE_STRING && nextCell.getStringCellValue().matches("^\\(.*\\)$")) {
                                nextCell.setCellValue("(".concat(inWords).concat(")"));
                            }
                        }
                    }
                }
                currentRowIndex++;
            }
        }

//5. Saving changes into file
        wb.write(new FileOutputStream(file));
        wb.close();
    }
}