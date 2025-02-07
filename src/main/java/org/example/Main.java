package org.example;

import java.sql.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.json.JSONArray;

import java.io.*;

public class Main {

    public static void main(String[] args) {
        String jdbcUrl = "jdbc:mysql://localhost:3306/data1";
        String username = "root";
        String password = "ecs123!@#";

        String excelFilePath = "C:\\Users\\SAMUEL JEBA DHAS.EINCAS1L-SJDH\\Downloads/BookData.xlsx";
        String jsonFilePath = "students_data.json";  // Path to save the JSON file

        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            int numberOfColumns = sheet.getRow(0).getPhysicalNumberOfCells();
            Row headerRow = sheet.getRow(0);

            List<Map<String, String>> studentDataList = new ArrayList<>();

            for (int i = 1; i <= sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, String> rowData = new LinkedHashMap<>();
                for (int j = 0; j < numberOfColumns; j++) {
                    Cell cell = row.getCell(j);
                    String cellValue = (cell != null) ? cell.toString() : "";
                    rowData.put(headerRow.getCell(j).toString(), cellValue);
                }


                insertDataIntoDatabase(rowData, jdbcUrl, username, password);

                // Add the row data to the list for JSON conversion
                studentDataList.add(rowData);
            }

            writeJsonToFile(studentDataList, jsonFilePath);

            workbook.close();
            fis.close();

            System.out.println("Excel file has been successfully converted to JSON and stored in MySQL!");

            Scanner scanner = new Scanner(System.in);
            System.out.print("Enter Admission Number or Name to search for a student: ");
            String searchQuery = scanner.nextLine();
            searchStudent(searchQuery, jdbcUrl, username, password);
            scanner.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Method to write the list of student data to a JSON file

    public static void writeJsonToFile(List<Map<String, String>> studentData, String filePath) throws IOException {
        FileWriter fileWriter = new FileWriter(filePath);
        JSONArray jsonArray = new JSONArray();

        for (Map<String, String> row : studentData) {
            // Create a new LinkedHashMap to preserve the order of insertion
            Map<String, Object> orderedRowData = new LinkedHashMap<>();

            orderedRowData.put("Name", row.get("Name"));
            orderedRowData.put("Admission No.", row.get("Admission No."));

            orderedRowData.put("Physics", row.get("Physics"));
            orderedRowData.put("Chemistry", row.get("Chemistry"));
            orderedRowData.put("Maths", row.get("Maths"));

            JSONObject jsonObject = new JSONObject(orderedRowData);
            jsonArray.put(jsonObject);
        }

        fileWriter.write(jsonArray.toString(4));
        fileWriter.close();
        System.out.println("Excel data has been converted to JSON and saved to " + filePath);
    }

    public static void insertDataIntoDatabase(Map<String, String> rowData, String jdbcUrl, String username, String password) {
        Connection connection = null;
        PreparedStatement preparedStatement = null;
        ResultSet resultSet = null;

        try {
            // Establish connection to MySQL
            connection = DriverManager.getConnection(jdbcUrl, username, password);

            String admissionNumberStr = rowData.get("Admission No.");
            if (admissionNumberStr == null || admissionNumberStr.trim().isEmpty()) {
                System.out.println("Admission Number is missing or empty. Skipping this row.");
                return;
            }

            int admissionNumber;
            try {
                admissionNumber = (int) Double.parseDouble(admissionNumberStr.trim());
            } catch (NumberFormatException e) {
                System.out.println("Invalid Admission Number format: " + admissionNumberStr);
                return;
            }

            // Check if the record already exists in the database
            String checkSql = "SELECT * FROM student WHERE admin = ?";
            preparedStatement = connection.prepareStatement(checkSql);
            preparedStatement.setInt(1, admissionNumber);
            resultSet = preparedStatement.executeQuery();

            if (resultSet.next()) {
                String updateSql = "UPDATE student SET name = ?, phy = ?, chem = ?, maths = ? WHERE admin = ?";
                preparedStatement = connection.prepareStatement(updateSql);
                preparedStatement.setString(1, rowData.get("Name"));

                int physicsMarks = parseMarks(rowData.get("Physics"));
                int chemistryMarks = parseMarks(rowData.get("Chemistry"));
                int mathsMarks = parseMarks(rowData.get("Maths"));

                preparedStatement.setInt(2, physicsMarks);
                preparedStatement.setInt(3, chemistryMarks);
                preparedStatement.setInt(4, mathsMarks);
                preparedStatement.setInt(5, admissionNumber);

                preparedStatement.executeUpdate();
            } else {
                String sql = "INSERT INTO student (admin, name, phy, chem, maths) VALUES (?, ?, ?, ?, ?)";
                preparedStatement = connection.prepareStatement(sql);
                preparedStatement.setInt(1, admissionNumber);
                preparedStatement.setString(2, rowData.get("Name"));

                int physicsMarks = parseMarks(rowData.get("Physics"));
                int chemistryMarks = parseMarks(rowData.get("Chemistry"));
                int mathsMarks = parseMarks(rowData.get("Maths"));

                preparedStatement.setInt(3, physicsMarks);
                preparedStatement.setInt(4, chemistryMarks);
                preparedStatement.setInt(5, mathsMarks);

                preparedStatement.executeUpdate();
            }

        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            try {
                if (resultSet != null) {
                    resultSet.close();
                }
                if (preparedStatement != null) {
                    preparedStatement.close();
                }
                if (connection != null) {
                    connection.close();
                }
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }

    public static int parseMarks(String marks) {
        if (marks == null || marks.trim().isEmpty()) {
            return 0;
        }

        try {
            double parsedMarks = Double.parseDouble(marks.trim());
            return (int) parsedMarks;
        } catch (NumberFormatException e) {
            System.out.println("Invalid marks format: " + marks);
            return 0;
        }
    }

    public static String[] getGradeAndGradePoint(int marks) {
        String grade;
        double gradePoint;

        if (marks > 90) {
            grade = "A1";
            gradePoint = 10.0;
        } else if (marks > 80) {
            grade = "A2";
            gradePoint = 9.0;
        } else if (marks > 70) {
            grade = "B1";
            gradePoint = 8.0;
        } else if (marks > 60) {
            grade = "B2";
            gradePoint = 7.0;
        } else if (marks > 50) {
            grade = "C1";
            gradePoint = 6.0;
        } else if (marks > 40) {
            grade = "C2";
            gradePoint = 5.0;
        } else if (marks >= 33) {
            grade = "D";
            gradePoint = 4.0;
        } else if (marks > 20) {
            grade = "E1";
            gradePoint = 0.0;
        } else {
            grade = "E2";
            gradePoint = 0.0;
        }

        return new String[]{grade, String.valueOf(gradePoint)};
    }

    public static void searchStudent(String searchQuery, String jdbcUrl, String username, String password) {
        Connection connection = null;
        PreparedStatement preparedStatement = null;
        ResultSet resultSet = null;

        try {
            connection = DriverManager.getConnection(jdbcUrl, username, password);

            String searchSql = "SELECT * FROM student WHERE admin = ? OR name = ?";
            preparedStatement = connection.prepareStatement(searchSql);
            preparedStatement.setString(1, searchQuery);
            preparedStatement.setString(2, searchQuery);

            resultSet = preparedStatement.executeQuery();

            if (resultSet.next()) {
                String name = resultSet.getString("name");
                int admissionNumber = resultSet.getInt("admin");
                int physicsMarks = resultSet.getInt("phy");
                int chemistryMarks = resultSet.getInt("chem");
                int mathsMarks = resultSet.getInt("maths");

                String[] physicsGrade = getGradeAndGradePoint(physicsMarks);
                String[] chemistryGrade = getGradeAndGradePoint(chemistryMarks);
                String[] mathsGrade = getGradeAndGradePoint(mathsMarks);

                double totalMarks = physicsMarks + chemistryMarks + mathsMarks;
                double percentage = (totalMarks / 300) * 100;

                // Print student details in original format
                System.out.println("Student found: {");
                System.out.println("  \"Name\" : \"" + name + "\",");
                System.out.println("  \"AdmissionNumber\" : \"" + admissionNumber + "\",");
                System.out.println("  \"Percentage\" : \"" + String.format("%.2f", percentage) + "\",");
                System.out.println("  \"Physics\" : \"" + physicsMarks + "\"");
                System.out.println("  \"Grade\" : \"" + physicsGrade[0] + "\", \"GradePoint\" : \"" + physicsGrade[1] + "\"");
                System.out.println("  \"Chemistry\" : \"" + chemistryMarks + "\"");
                System.out.println("  \"Grade\" : \"" + chemistryGrade[0] + "\", \"GradePoint\" : \"" + chemistryGrade[1] + "\"");
                System.out.println("  \"Maths\" : \"" + mathsMarks + "\"");
                System.out.println("  \"Grade\" : \"" + mathsGrade[0] + "\", \"GradePoint\" : \"" + mathsGrade[1] + "\"");
                System.out.println("}");
            } else {
                System.out.println("No student found with the given Admission Number or Name.");
            }

        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            try {
                if (resultSet != null) {
                    resultSet.close();
                }
                if (preparedStatement != null) {
                    preparedStatement.close();
                }
                if (connection != null) {
                    connection.close();
                }
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }
}