package ru.nikiton;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.omg.CosNaming.NamingContextExtPackage.StringNameHelper;

import java.io.*;
import java.util.*;

public class App {

    public static void main(String[] args){
        try{
            PrintWriter logFile = new PrintWriter(new FileOutputStream("ExcelWorker_log.txt"));
            File currentFolder = new File(".");
            File[] filesInFolder;
            DirFilter directoryFilter = new DirFilter(".*\\.xls$");
            filesInFolder = currentFolder.listFiles(directoryFilter);
            logFile.println("Запуск программы: " + new Date());
            logFile.println("-----------------------------------------------");
            logFile.println("В папке найдено .xls файлов: " + filesInFolder.length);
            int countFiles = 0;
            for(File xlsFile : filesInFolder){
                countFiles++;
                logFile.println("\nРабота с файлом №" + countFiles + " : \"" + xlsFile.getName() + "\"");
                HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(xlsFile));
                //1. 	Parsing .xls file to determine file type:
                int countWbSheets = wb.getNumberOfSheets();
                boolean operatorsSheetFound = false;
                boolean contractorsSheetFound = false;
                boolean errorFound = false;
                for(int i=0; i < countWbSheets; i++){
                    if(wb.getSheetAt(i).getSheetName().contains("Операторы") || wb.getSheetAt(i).getSheetName().contains("Универсалы")){
                        logFile.println("Найден Лист " + wb.getSheetAt(i).getSheetName());
                        if(!operatorsSheetFound)
                            operatorsSheetFound = true;
                        else{
                            logFile.println("!!! ОШИБКА при работе с файлом: \"" + xlsFile.getName() + "\"");
                            logFile.println("В файле присутствуют дублирующие Листы Операторов/Универсалов!");
                            errorFound = true;
                            break;
                        }
                    }
                    if(wb.getSheetAt(i).getSheetName().contains("Акт")){
                        logFile.println("Найден Лист " + wb.getSheetAt(i).getSheetName());
                        contractorsSheetFound = true;
                    }
                }
                wb.close();
                if(errorFound)
                    continue;
                if(operatorsSheetFound && contractorsSheetFound){
                    logFile.println("!!! ОШИБКА при работе с файлом: \"" + xlsFile.getName() + "\"");
                    logFile.println("В файле присутствуют Листы как Операторов/Универсалов, так и Подрядчиков!");
                    continue;
                }
                if(!operatorsSheetFound && !contractorsSheetFound){
                    logFile.println("!!! ОШИБКА при работе с файлом: \"" + xlsFile.getName() + "\"");
                    logFile.println("В файле отствуют Листы Операторов/Универсалов или Подрядчиков!");
                    continue;
                }
                if(operatorsSheetFound){
                    logFile.println("Работаем с Операторами/Универсалами");
                    try {
                        //Enter point
                        ArrayList<String> warnings = OperatorWorker.make(xlsFile);
                        if(warnings.isEmpty())
                            logFile.println("Файл отработан успешно!");
                        else {
                            logFile.println("Файл отработан с замечаниями:");
                            for(String warning : warnings) {
                                logFile.println(warning);
                            }
                        }
                    }
                    catch (ExcelWorkerException e){
                        logFile.print("!!!ОШИБКА!!! Неверная структура файла! " + e.toString() + "\n");
                    }
                    catch(Exception e){
                        logFile.println("!!!ОШИБКА!!! Системная:");
                        e.printStackTrace(logFile);
                    }
                }
                else{
                    logFile.println("Работаем с Подрядчиками");
                    try {
                        //Enter point
                        ContractorWorker.make(xlsFile);
                        logFile.println("Файл отработан успешно!");
                    }
                    catch (ExcelWorkerException e){
                        logFile.print("!!!ОШИБКА!!! Неверная структура файла! " + e.toString() + "\n");
                    }
                    catch(Exception e){
                        logFile.println("!!!ОШИБКА!!! Системная: \n");
                        e.printStackTrace(logFile);
                    }
                }
            }
            //Closing Log stream in file
            logFile.close();
        }
        catch(Exception e){
            try {
                FileOutputStream errorFile = new FileOutputStream("ExcelWorker_error.txt");
                PrintStream backup = System.err;
                PrintStream ourStr = new PrintStream(errorFile);
                System.setErr(ourStr);
                e.printStackTrace();
                ourStr.close();
                errorFile.close();
                System.setErr(backup);
            } catch(IOException e1){
                e1.printStackTrace();
            }
        }
    }
}