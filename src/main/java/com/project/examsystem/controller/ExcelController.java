package com.project.examsystem.controller;

import com.project.examsystem.exception.FileException;
import com.project.examsystem.service.impl.ExcelServiceImpl;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;


@RestController
@RequestMapping("/api")
@RequiredArgsConstructor
public class ExcelController {
    private final ExcelServiceImpl excelService;

    @PostMapping("/excel")
    public void excelReader(@RequestParam("file") MultipartFile excel) {
        try {
            InputStream inputStream = excel.getInputStream();
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            List<String> answerKey = new ArrayList<>();
            List<Double> scores = new ArrayList<>();
            Map<String, List<String>> studentAnswers = new HashMap<>();
            Map<String, Double> studentScores = new HashMap<>();
            Map<String, Double> studentScores2 = new HashMap<>();
            Map<Integer, Double> subgroupAnalysis = new HashMap<>();
            Map<Integer, Double> itemDifficultyIndex = new HashMap<>();
            Map<Integer, Double> itemVariances = new HashMap<>();
            Map<Integer, Double> itemStandardDeviations = new HashMap<>();
            Map<String, Double> percentageMap = new HashMap<>();
            Map<Integer, Double> itemReliabilityCoefficient = new HashMap<>();
            int numberOfQuestions = (sheet.getRow(0).getPhysicalNumberOfCells() - 2);
            int numberOfStudents, aCount, bCount, cCount, dCount, eCount, emptyCount;
            double coefficientOfSkewness, kr20, kr21, averageDifficultyCoefficient, totalScore, coefficientOfKurtosis,
                    standardErrorCoefficientOfTheTest, questionScore = 100 / (double) numberOfQuestions, average = 0, average2 = 0, maxValue, minValue, range, median, relativeCoefficientOfVariation,
                    standardDeviation, variance;
            Map<String, Map<String, Integer>> optionCountsPerQuestion = new HashMap<>();
            HashMap<Integer, List<String>> answersToQuestions = new HashMap<>();
            Row answerKeyRow = sheet.getRow(1);
            for (int j = 2; j < numberOfQuestions + 2; j++) {
                answerKey.add(answerKeyRow.getCell(j).getStringCellValue());
            }
            for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                /*double cellValue = row.getCell(1).getNumericCellValue(); // 1. column has student number
                String studentNumber = String.valueOf((long) cellValue);*/
                String studentNumber = row.getCell(1).getStringCellValue();
                List<String> answers = studentAnswers.get(studentNumber);
                if (answers == null) {
                    answers = new ArrayList<>();
                    for (int j = 2; j < numberOfQuestions + 2; j++) {
                        if (excelService.isCellEmpty(row.getCell(j))) {
                            answers.add("Boş");
                        } else {
                            answers.add(String.valueOf(row.getCell(j)));
                        }
                    }
                    studentAnswers.put(studentNumber, answers);
                }
            }
            int startRow = 2;
            int endRow = sheet.getLastRowNum();
            int startColumn = 2;
            int endColumn = numberOfQuestions + 2;
            for (int column = startColumn; column < endColumn; column++) {
                List<String> dataList = new ArrayList<>();
                for (int row = startRow; row <= endRow; row++) {
                    if (!excelService.isCellEmpty(sheet.getRow(row).getCell(column))) {
                        String cellValue = sheet.getRow(row).getCell(column).getStringCellValue();
                        dataList.add(cellValue);
                    } else {
                        dataList.add("Boş");
                    }
                }
                int key = column - startColumn + 1;
                answersToQuestions.put(key, dataList);
            }
            for (Map.Entry<Integer, List<String>> entry : answersToQuestions.entrySet()) {
                aCount = 0;
                bCount = 0;
                cCount = 0;
                dCount = 0;
                eCount = 0;
                emptyCount = 0;
                HashMap<String, Integer> h = new HashMap<>();
                Integer key = entry.getKey();
                List<String> values = entry.getValue();
                for (String s : values) {
                    if (s.equals("A")) {
                        aCount++;
                    } else if (s.equals("B")) {
                        bCount++;
                    } else if (s.equals("C")) {
                        cCount++;
                    } else if (s.equals("D")) {
                        dCount++;
                    } else if (s.equals("E")) {
                        eCount++;
                    } else {
                        emptyCount++;
                    }
                }
                h.put("A", aCount);
                h.put("B", bCount);
                h.put("C", cCount);
                h.put("D", dCount);
                h.put("E", eCount);
                h.put("Boş", emptyCount);
                optionCountsPerQuestion.put(key + ".Soru", h);
            }
            for (Map.Entry<String, Map<String, Integer>> entry : optionCountsPerQuestion.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            System.out.println("-- Students' answers --");
            for (Map.Entry<String, List<String>> entry : studentAnswers.entrySet()) {
                totalScore = 0;
                double totalScore2 = 0;
                String anahtar = entry.getKey();
                List<String> degerler = entry.getValue();
                for (int i = 0; i < answerKey.size(); i++) {
                    if (answerKey.get(i).equals(degerler.get(i))) {
                        totalScore += questionScore;
                        totalScore2 += 1;
                    }
                }
                studentScores.put(anahtar, totalScore);
                studentScores2.put(anahtar, totalScore2);
                scores.add(totalScore2);
                System.out.println(anahtar + " -> " + degerler.toString());
            }
            numberOfStudents = studentScores.size();
            System.out.println("-- Öğrencilerin almış olduğu puanlar / Yuzdelik--");
            for (Map.Entry<String, Double> entry : studentScores.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            System.out.println("-- Öğrencilerin almış olduğu puanlar / 1-0 --");
            for (Map.Entry<String, Double> entry : studentScores2.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < answerKey.size(); i++) {
                sb.setLength(0);
                sb.append(i + 1).append(".Soru");
                double percentage = optionCountsPerQuestion.get(sb.toString()).get(answerKey.get(i));
                percentage = ((100 * percentage) / numberOfStudents) / 100;
                percentageMap.put(sb.toString(), percentage);
            }
            System.out.println("Sorularin Dogru Cevaplanma Yuzdeleri:");
            for (Map.Entry<String, Double> entry : percentageMap.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            maxValue = Collections.max(scores);
            minValue = Collections.min(scores);
            range = maxValue - minValue;
            average = excelService.calculateAverage(studentScores, numberOfStudents);
            average2 = excelService.calculateAverage(studentScores2, numberOfStudents);
            standardDeviation = excelService.calculateStandardDeviation(scores, average2);
            variance = Math.pow(standardDeviation, 2);
            median = excelService.calculateMedian(scores);
            relativeCoefficientOfVariation = excelService.calculateRelativeCoefficientOfVariation(standardDeviation, average2);
            List<Double> modes = excelService.calculateMode(scores);
            coefficientOfSkewness = excelService.calculateCoefficientOfSkewness(average2, standardDeviation, median);
            kr20 = excelService.calculateKR20(percentageMap, variance, numberOfQuestions);
            kr21 = excelService.calculateKR21(average2, variance, numberOfQuestions);
            coefficientOfKurtosis = excelService.calculateCoefficientOfKurtosis(scores, average2, standardDeviation, numberOfStudents);
            averageDifficultyCoefficient = excelService.calculateAverageDifficultyCoefficient(average2, numberOfQuestions);
            standardErrorCoefficientOfTheTest = excelService.calculateStandardErrorCoefficientOfTheTest(kr20, standardDeviation);
            for (int i = 0; i < numberOfQuestions; i++) {
                double indexOfDistinctiveness = excelService.subGroupAndSuperGroupAnalysis(studentScores, studentAnswers, i, answerKey, numberOfStudents, numberOfQuestions);
                subgroupAnalysis.put(i + 1, indexOfDistinctiveness);
            }
            for (int i = 0; i < numberOfQuestions; i++) {
                double difficultyIndex = excelService.calculateItemDifficultyIndex(studentScores, studentAnswers, i, answerKey, numberOfStudents, numberOfQuestions);
                itemDifficultyIndex.put(i + 1, difficultyIndex);
            }
            for (int i = 0; i < numberOfQuestions; i++) {
                double itemVariance = excelService.calculateItemVariance(percentageMap, itemDifficultyIndex, i);
                itemStandardDeviations.put(i + 1, Math.sqrt(itemVariance));
                itemVariances.put(i + 1, itemVariance);
            }
            for (int i = 0; i < numberOfQuestions; i++) {
                double itemVariance = excelService.calculateItemReliabilityCoefficient(subgroupAnalysis, itemStandardDeviations, i);
                itemReliabilityCoefficient.put(i + 1, itemVariance);
            }
            System.out.println("Index of Distinctiveness:");
            for (Map.Entry<Integer, Double> entry : subgroupAnalysis.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            System.out.println("Item Difficulty Index:");
            for (Map.Entry<Integer, Double> entry : itemDifficultyIndex.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            System.out.println("Item Variance:");
            for (Map.Entry<Integer, Double> entry : itemVariances.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            System.out.println("Item Standart Deviation:");
            for (Map.Entry<Integer, Double> entry : itemStandardDeviations.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            System.out.println("Item Reliability Coefficient:");
            for (Map.Entry<Integer, Double> entry : itemReliabilityCoefficient.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            System.out.println("-- RESULTS --");
            System.out.println("Average: " + average2 + " - %" + average +
                    "\nMaximum: " + maxValue + " - %" + maxValue * questionScore +
                    "\nMinimum: " + minValue + " - %" + minValue * questionScore +
                    "\nRange: " + range +
                    "\nStandard Deviation: " + standardDeviation +
                    "\nVariance: " + variance +
                    "\nMedian: " + median +
                    "\nRelative Coefficient of Variation: " + relativeCoefficientOfVariation +
                    "\nMode: " + modes +
                    "\nCoefficient of Skewness: " + coefficientOfSkewness +
                    "\nCoefficient of Kurtosis: " + coefficientOfKurtosis +
                    "\nKR20: " + kr20 +
                    "\nKR21: " + kr21 +
                    "\nStandard Error Coefficient of the test: " + standardErrorCoefficientOfTheTest +
                    "\nAverage Difficulty Coefficient of the test: " + averageDifficultyCoefficient
            );
            Map<String, Double> zPoints;
            zPoints = excelService.calculateZPoints(studentScores2, standardDeviation, average2);
            System.out.println("Z Points: ");
            for (Map.Entry<String, Double> entry : zPoints.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            System.out.println("T Points: ");
            Map<String, Double> tPoints;
            tPoints = excelService.calculateTPoints(zPoints);
            for (Map.Entry<String, Double> entry : tPoints.entrySet()) {
                System.out.println(entry.getKey() + " -> " + entry.getValue());
            }
            workbook.close();
            inputStream.close();
        } catch (IOException e) {
            throw new FileException("Dosya okunurken hata oluştu");
        }
    }
}
