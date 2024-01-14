# Internship Task: Excel Analysis

## Task Description

This repository contains my solution to the internship task for analyzing Excel sheets. The task includes writing a Java program to programmatically analyze an Excel file based on specific criteria.

## Instructions

1. Clone the repository locally.
2. Open the project in Eclipse or any Java IDE.
3. Run the `ExcelAnalysis` class to execute the analysis.
4. View the console output for the analysis results.

## Code Structure

- `ExcelAnalysis.java`: Main Java file containing the code for Excel analysis.
- `output.txt`: Output file containing the console output during analysis.

## Sample Excel Data

The Excel sheet used for testing includes the following columns:

- Position ID
- Position Status
- Time
- Time Out
- Timecard Hours (as Time)
- Pay Cycle Start Date
- Pay Cycle End Date
- Employee Name
- File Number

## Analysis Criteria

The program analyzes the data and prints information about employees who:
- Have worked for 7 consecutive days.
- Have less than 10 hours between shifts but greater than 1 hour.
- Have worked for more than 14 hours in a single shift.

