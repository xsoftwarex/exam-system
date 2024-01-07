package com.project.examsystem.service.impl;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.springframework.stereotype.Service;

import java.util.*;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl {

    public double calculateAverage(Map<String, Double> studentScores, int numberOfStudents) {
        double scoreCount = 0;
        for (Map.Entry<String, Double> entry : studentScores.entrySet()) {
            scoreCount += entry.getValue();
        }
        return scoreCount / numberOfStudents;
    }

    public double calculateStandardDeviation(List<Double> scoreList, double average) {
        double count = 0;
        for (Double value : scoreList) {
            value -= average;
            count += Math.pow(value, 2);
        }
        return Math.sqrt(count / (scoreList.size()));
    }

    public double calculateMedian(List<Double> scoreList) {
        List<Double> sortedList = new ArrayList<>(scoreList);
        Collections.sort(sortedList);
        int size = sortedList.size();
        double median;
        if (size % 2 == 0) {
            median = (sortedList.get(size / 2) + sortedList.get(size / 2 - 1)) / 2.0;
        } else {
            median = sortedList.get(size / 2);
        }
        return median;
    }

    public List<Double> calculateMode(List<Double> list) {
        Map<Double, Integer> frequencyMap = new HashMap<>();
        int maxFrequency = 0;
        for (double num : list) {
            int frequency = frequencyMap.getOrDefault(num, 0) + 1;
            frequencyMap.put(num, frequency);
            maxFrequency = Math.max(maxFrequency, frequency);
        }
        List<Double> modes = new ArrayList<>();
        for (Map.Entry<Double, Integer> entry : frequencyMap.entrySet()) {
            double num = entry.getKey();
            int frequency = entry.getValue();

            if (frequency == maxFrequency) {
                modes.add(num);
            }
        }
        return modes;
    }

    public double calculateCoefficientOfKurtosis(List<Double> scores, double average, double standardDeviation, int numberOfStudents) {
        double total = 0;
        for (Double score : scores) {
            total = total + Math.pow((score - average), 4);
        }
        return ((total / (numberOfStudents * Math.pow(standardDeviation, 4))) - 3);
    }

    public double calculateAverageDifficultyCoefficient(double average, int numberOfQuestions) {
        return average / (numberOfQuestions);
    }

    public double calculateKR20(Map<String, Double> percentageMap, double variance, int numberOfQuestions) {
        double kr20, total = 0;
        for (Map.Entry<String, Double> entry : percentageMap.entrySet()) {
            total = total + ((entry.getValue()) * (1 - (entry.getValue())));
        }
        kr20 = (((double) numberOfQuestions) / (((double) numberOfQuestions) - 1)) * (1 - (total / variance));
        return kr20;
    }

    public double calculateKR21(double average, double variance, int numberOfQuestions) {
        double kr21 = numberOfQuestions * (1 - (((numberOfQuestions * average) - Math.pow(average, 2)) / (numberOfQuestions * variance)));
        kr21 = kr21 / (numberOfQuestions - 1);
        return kr21;
    }

    public Map<String, Double> calculateZPoints(Map<String, Double> studentScores, double standardDeviation, double average) {
        Map<String, Double> zPoints = new HashMap<>();
        for (Map.Entry<String, Double> entry : studentScores.entrySet()) {
            double zPoint = (entry.getValue() - average) / standardDeviation;
            zPoints.put(entry.getKey(), zPoint);
        }
        return zPoints;
    }

    public Map<String, Double> calculateTPoints(Map<String, Double> zPoints) {
        Map<String, Double> tPoints = new HashMap<>();
        for (Map.Entry<String, Double> entry : zPoints.entrySet()) {
            double tPoint = (10 * entry.getValue()) + 50;
            tPoints.put(entry.getKey(), tPoint);
        }
        return tPoints;
    }

    public boolean isAnsweredCorrectByStudent(String studentName, Map<String, List<String>> studentAnswers, int questionNumber, List<String> answerKey) {
        List<String> answers = studentAnswers.get(studentName);
        return answers.get(questionNumber).equals(answerKey.get(questionNumber));
    }

    public LinkedHashMap<String, Double> sortMap(Map<String, Double> studentScores) {
        return studentScores.entrySet()
                .stream()
                .sorted(Map.Entry.<String, Double>comparingByValue().reversed())
                .collect(
                        LinkedHashMap::new,
                        (map, entry) -> map.put(entry.getKey(), entry.getValue()),
                        LinkedHashMap::putAll
                );
    }
    // TODO: Asagidaki metotlarda kod tekrari vardir.
    // TODO: Duzenlenecek.

    public double subGroupAndSuperGroupAnalysis(Map<String, Double> studentScores, Map<String, List<String>> studentAnswers, int questionNumber, List<String> answerKey, int numberOfStudents, int numberOfQuestions) {
        LinkedHashMap<String, Double> sortedStudentScores = sortMap(studentScores);
        int number = (int) Math.round(numberOfStudents * (0.27));
        int alt = 0, ust = 0;
        double indexOfDistinctiveness;
        int count = 0;
        for (Map.Entry<String, Double> entry : sortedStudentScores.entrySet()) {
            if (count < number) {
                String key = entry.getKey();
                if (isAnsweredCorrectByStudent(key, studentAnswers, questionNumber, answerKey)) {
                    ust = ust + 1;
                }
            } else {
                break;
            }
            count++;
        }
        count = 0;
        Iterator<Map.Entry<String, Double>> iterator = sortedStudentScores.entrySet().iterator();
        while (iterator.hasNext()) {
            Map.Entry<String, Double> entry = iterator.next();
            if (count >= sortedStudentScores.size() - number) {
                String key = entry.getKey();
                //Double value = entry.getValue();
                //System.out.println("Key: " + key + ", Value: " + value);
                if (isAnsweredCorrectByStudent(key, studentAnswers, questionNumber, answerKey)) {
                    alt = alt + 1;
                }
            }
            count++;
        }
        indexOfDistinctiveness = (double) (ust - alt) / number;
        return indexOfDistinctiveness;
    }

    public double calculateItemDifficultyIndex(Map<String, Double> studentScores, Map<String, List<String>> studentAnswers, int questionNumber, List<String> answerKey, int numberOfStudents, int numberOfQuestions) {
        LinkedHashMap<String, Double> sortedStudentScores = sortMap(studentScores);
        int number = (int) Math.round(numberOfStudents * (0.27));
        int alt = 0, ust = 0;
        double itemDifficultyIndex;
        int count = 0;
        for (Map.Entry<String, Double> entry : sortedStudentScores.entrySet()) {
            if (count < number) {
                String key = entry.getKey();
                if (isAnsweredCorrectByStudent(key, studentAnswers, questionNumber, answerKey)) {
                    ust = ust + 1;
                }
            } else {
                break;
            }
            count++;
        }
        count = 0;
        Iterator<Map.Entry<String, Double>> iterator = sortedStudentScores.entrySet().iterator();
        while (iterator.hasNext()) {
            Map.Entry<String, Double> entry = iterator.next();
            if (count >= sortedStudentScores.size() - number) {
                String key = entry.getKey();
                //Double value = entry.getValue();
                //System.out.println("Key: " + key + ", Value: " + value);
                if (isAnsweredCorrectByStudent(key, studentAnswers, questionNumber, answerKey)) {
                    alt = alt + 1;
                }
            }
            count++;
        }
        itemDifficultyIndex = (double) (ust + alt) / (number * 2);
        return itemDifficultyIndex;
    }
    public double calculateItemVariance(Map<String, Double> percentageMap, Map<Integer, Double> itemDifficultyIndex, int index) {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append(index + 1);
        stringBuilder.append(".Soru");
        String result = stringBuilder.toString();
        double wrongAnswerPercentage = 1 - (percentageMap.get(result));
        double difficultyIndex = itemDifficultyIndex.get(index + 1);
        return wrongAnswerPercentage * difficultyIndex;
    }
    public double calculateItemReliabilityCoefficient(Map<Integer, Double> subGroupAnalysis, Map<Integer, Double> itemStandardDeviations, int index){
        double indexOfDistinctiveness = subGroupAnalysis.get(index + 1);
        double itemStandardDeviation = itemStandardDeviations.get(index + 1);
        return indexOfDistinctiveness * itemStandardDeviation;
    }

    public double calculateStandardErrorCoefficientOfTheTest(double kr20, double standardDeviation) {
        return standardDeviation * (Math.sqrt(1 - kr20));
    }
    public double calculateRelativeCoefficientOfVariation(double standardDeviation, double average){
        return (standardDeviation / average) * 100;
    }
    public double calculateCoefficientOfSkewness(double average, double standardDeviation, double median){
        return (3 * (average - median)) / standardDeviation;
    }

    public static boolean isCellEmpty(final Cell cell) {
        if (cell == null) { // use row.getCell(x, Row.CREATE_NULL_AS_BLANK) to avoid null cells
            return true;
        }

        if (cell.getCellType() == CellType.BLANK) {
            return true;
        }

        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) {
            return true;
        }

        return false;
    }
}

