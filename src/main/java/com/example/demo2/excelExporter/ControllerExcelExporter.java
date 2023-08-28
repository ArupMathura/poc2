package com.example.demo2.excelExporter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import com.example.demo2.controller.HelloController;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class ControllerExcelExporter {

    public void exportControllersToExcel(List<Class<HelloController>> controllerClasses, String outputPath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Controller Info");

        int rowIndex = 0;
        Row headerRow = sheet.createRow(rowIndex++);
        headerRow.createCell(0).setCellValue("Controller");
        headerRow.createCell(1).setCellValue("RequestMapping");
        headerRow.createCell(2).setCellValue("RequestMethod");
        headerRow.createCell(3).setCellValue("Method");

        for (Class<?> controllerClass : controllerClasses) {
            RestController restControllerAnnotation = controllerClass.getAnnotation(RestController.class);
            if (restControllerAnnotation != null) {
                RequestMapping classRequestMapping = controllerClass.getAnnotation(RequestMapping.class);
                String controllerRequestMapping = "";
                RequestMethod[] classRequestMethods = null;

                if (classRequestMapping != null) {
                    controllerRequestMapping = classRequestMapping.value()[0];
                    classRequestMethods = classRequestMapping.method();
                }

                Method[] methods = controllerClass.getDeclaredMethods();
                for (Method method : methods) {
                    RequestMapping methodRequestMapping = method.getAnnotation(RequestMapping.class);
                    if (methodRequestMapping != null) {
                        String methodRequestMappingValue = methodRequestMapping.value()[0];
                        RequestMethod[] methodRequestMethods = methodRequestMapping.method();

                        Row dataRow = sheet.createRow(rowIndex++);
                        dataRow.createCell(0).setCellValue(controllerClass.getSimpleName());
                        dataRow.createCell(1).setCellValue(controllerRequestMapping + methodRequestMappingValue);

                        if (methodRequestMethods != null && methodRequestMethods.length > 0) {
                            dataRow.createCell(2).setCellValue(methodRequestMethods[0].toString());
                        }

                        dataRow.createCell(3).setCellValue(method.getName());
                    }
                }
            }
        }

        try (FileOutputStream outputStream = new FileOutputStream(outputPath)) {
            workbook.write(outputStream);
        }
    }

    public static void main(String[] args) {
        ControllerExcelExporter exporter = new ControllerExcelExporter();
        List<Class<HelloController>> controllerClasses = getControllerClasses("com.example.demo2.controller"); // Adjust this package
        String outputPath = "D:\\Eclipse-Workspace\\demo2\\excel\\ControllerInfo.xlsx";

        try {
            exporter.exportControllersToExcel(controllerClasses, outputPath);
            System.out.println("Data exported successfully.");
        } catch (IOException e) {
            System.err.println("Error exporting data: " + e.getMessage());
        }
    }

    private static List<Class<HelloController>> getControllerClasses(String basePackage) {
        // This method should collect controller classes using the approach shown in the previous responses.
        // For simplicity, you can return a predefined list of controller classes here.
        List<Class<HelloController>> controllerClasses = new ArrayList<>();
        // Add your controller classes to the list
        controllerClasses.add(HelloController.class); // Add your controller class here
        return controllerClasses;
    }
}

