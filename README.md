# Excel-and-Python-Combination-for-Document-Generation
Using Excel and Python for semi-automatic document creation.

BEFORE USING THE RECOMMENDATION GENERATOR
1. Installing Python 3
2. Installing packages:
  * pandas
  * os
  * python-docx
  * openpyxl
3. Setting the following paths:
  * working directory, e.g. "D:\\_PROJECTS\\excel_python_document" (in the Python script)
  * Python script , e.g. Args = """D:\\_PROJECTS\\excel_python_document\\recommendationGenerator.py""" (in the Excel macro)
  * Python .exe file, e.g. "C:\\Users\\ludek\\Programs\\Python\\Python310\\python.exe" (in the Excel macro)

HOW TO USE RECOMMENDATION GENERATOR
1. Open the recommendationGenerator.xlsx file.
2. Select "Yes" in the "PRESENT" column for those diagnoses for which you want to generate corresponding recommendations and save the changes.
3. After selecting diagnosis, push the button "SAVE TEXT TO A WORD DOCUMENT".
4. Word document "recommendations.docs" will be generated with texts for all selected diagnoses. You will find it in the same folder as the Excel file.
5. It is possible to add new rows with new diagnoses and new texts (with limitation of 20 text columns).

===========================================

Alternative Excel macro for older versions of Excel (e.g. Excel 2016)

Sub python_script()
 
' link_python_excel Macro
' Declare all variables
Dim objShell As Object
Dim PythonExe, PythonScript As String
     
    'Create a new Shell Object
    Set objShell = VBA.CreateObject("Wscript.shell")
         
    'Provide the file path to the Python Exe
    PythonExe = """C:\\Users\\ludek\\Programs\\Python\\Python310\\python.exe"""
         
    'Provide the file path to the Python script
    PythonScript = "D:\\_PROJECTS\\excel_python_document\\recommendationGenerator.py"""
         
    'Run the Python script
    objShell.Run PythonExe & PythonScript
     
End Sub
