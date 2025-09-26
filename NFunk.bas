Attribute VB_Name = "NFunk"
Option Compare Database
Option Explicit


Public Const TableName As String = "Студенты отличники"

Public Const Field_Num As String = "НомерЗачКнижки"
Public Const Field_FIO As String = "ФИО"
Public Const Field_Subject As String = "НаименованиеДисциплины"
Public Const Field_Grade As String = "Оценка"
Public Const Field_Kafedra As String = "Кафедра"


Public Const TableName2 As String = "Зачетная ведомость"
Public Const GS_Stud As String = "№ зач_книжки"
Public Const GS_Subj As String = "№ дисциплины"
Public Const GS_Teacher As String = "№ преподавателя"
Public Const GS_Pass As String = "Зачтено/Незачтено"
Public Const GS_Grade As String = "Оценка"
Public Const GS_Date As String = "Дата сдачи"


Sub N1()
    CreateTable
End Sub

Sub N2()
    DeleteTable
End Sub

Sub N3()
    AddField_Kafedra
End Sub

Sub N4()
    DeleteField_Kafedra
End Sub

Sub N5()
    CreateGradeSheetTable
End Sub
Sub N6()
    DeleteGradeSheetTable
End Sub
