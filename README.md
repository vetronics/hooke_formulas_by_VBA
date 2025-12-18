# Hooke's Law Excel VBA Calculator

This Excel workbook contains VBA macros to calculate values based on Hooke's Law formula.

## Features

- Interactive checkbox-based forms for three different calculations
- Calculate Force from spring constant and extension
- Calculate Extension from force and spring constant
- Calculate Spring Constant from force and extension
- Input validation to prevent division by zero

## Hooke's Law Formula

**F = -k × ΔS**

Where:
- **F**: Force
- **k**: Spring constant
- **ΔS**: Extension

## How to Use

1. Open the Excel file (.xlsm format)
2. Enable macros when prompted
3. The form has three checkboxes:
   - **Force Checkbox**: Enter the spring constant and extension values to calculate force
   - **Extension Checkbox**: Enter the force and spring constant values to calculate extension
   - **Spring Constant Checkbox**: Enter the force and extension values to calculate spring constant
4. Input your values in the dialog boxes that appear
5. The result will be displayed in a message box

## Macros Included

### box_hooke_Click()
Calculates force using the formula: F = |−k × ΔS|

### exstention_box_Click()
Calculates extension using the formula: x = F / k

### konstant_box_Click()
Calculates spring constant using the formula: k = F / x

## Requirements

- Microsoft Excel with macro support enabled (.xlsm file)