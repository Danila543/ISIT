Attribute VB_Name = "NFunk"
Option Compare Database
Option Explicit


Public Const TableName As String = "�������� ���������"

Public Const Field_Num As String = "��������������"
Public Const Field_FIO As String = "���"
Public Const Field_Subject As String = "����������������������"
Public Const Field_Grade As String = "������"
Public Const Field_Kafedra As String = "�������"


Public Const TableName2 As String = "�������� ���������"
Public Const GS_Stud As String = "� ���_������"
Public Const GS_Subj As String = "� ����������"
Public Const GS_Teacher As String = "� �������������"
Public Const GS_Pass As String = "�������/���������"
Public Const GS_Grade As String = "������"
Public Const GS_Date As String = "���� �����"


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
