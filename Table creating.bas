Attribute VB_Name = "Table creating"
Option Compare Database
Option Explicit

Sub CreateTable()
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim pk  As DAO.Index

    Set dbs = CurrentDb

    DeleteTable

    Set tdf = dbs.CreateTableDef(TableName)


    With tdf.Fields
        .Append tdf.CreateField(Field_Num, dbLong)
        .Append tdf.CreateField(Field_FIO, dbText, 100)
        .Append tdf.CreateField(Field_Subject, dbText, 100)
        .Append tdf.CreateField(Field_Grade, dbInteger)
    End With


    Set pk = tdf.CreateIndex("PK_" & TableName)
    pk.Primary = True
    pk.Unique = True
    pk.Fields.Append pk.CreateField(Field_Num)
    tdf.Indexes.Append pk


    dbs.TableDefs.Append tdf
    dbs.TableDefs.Refresh

    Set tdf = Nothing
    Set dbs = Nothing
End Sub


Sub DeleteTable()
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef

    Set dbs = CurrentDb

    For Each tdf In dbs.TableDefs
        If tdf.Name = TableName Then
            dbs.TableDefs.Delete TableName
            dbs.TableDefs.Refresh
            Exit For
        End If
    Next tdf

    Set dbs = Nothing
End Sub

Sub AddField_Kafedra()
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim t As DAO.TableDef

    Set dbs = CurrentDb

   
    DeleteField_Kafedra

    For Each t In dbs.TableDefs
        If t.Name = TableName Then
            Set tdf = t
            Exit For
        End If
    Next t
    If tdf Is Nothing Then Exit Sub


    Set fld = tdf.CreateField(Field_Kafedra, dbText, 50)
    tdf.Fields.Append fld
    tdf.Fields.Refresh

    Set tdf = Nothing
    Set dbs = Nothing
End Sub

Sub DeleteField_Kafedra()
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim t As DAO.TableDef
    Dim fld As DAO.Field

    Set dbs = CurrentDb


    For Each t In dbs.TableDefs
        If t.Name = TableName Then
            Set tdf = t
            Exit For
        End If
    Next t
    If tdf Is Nothing Then Exit Sub


    For Each fld In tdf.Fields
        If fld.Name = Field_Kafedra Then
            tdf.Fields.Delete Field_Kafedra
            tdf.Fields.Refresh
            Exit For
        End If
    Next fld

    Set tdf = Nothing
    Set dbs = Nothing
End Sub


Sub CreateGradeSheetTable()
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim t As DAO.TableDef
    Dim pk As DAO.Index

    Set db = CurrentDb


    For Each t In db.TableDefs
        If t.Name = TableName2 Then
            db.TableDefs.Delete TableName2
            Exit For
        End If
    Next t
    

    Set td = db.CreateTableDef(TableName2)

    With td.Fields
        .Append td.CreateField(GS_Stud, dbInteger)
        .Append td.CreateField(GS_Subj, dbInteger)
        .Append td.CreateField(GS_Teacher, dbInteger)
        .Append td.CreateField(GS_Pass, dbText, 20)
        .Append td.CreateField(GS_Grade, dbInteger)
        .Append td.CreateField(GS_Date, dbDate)
    End With


    td.Fields(GS_Pass).ValidationRule = "In (""Зачтено"",""Незачтено"")"
    td.Fields(GS_Pass).ValidationText = "Допустимы только: Зачтено / Незачтено."
    td.Fields(GS_Grade).ValidationRule = "Between 1 And 5"
    td.Fields(GS_Grade).ValidationText = "Оценка должна быть от 1 до 5."


    Set pk = td.CreateIndex("PK_" & TableName2)
    pk.Primary = True
    pk.Unique = True
    pk.Fields.Append pk.CreateField(GS_Stud)
    pk.Fields.Append pk.CreateField(GS_Subj)
    pk.Fields.Append pk.CreateField(GS_Teacher)
    td.Indexes.Append pk


    db.TableDefs.Append td
End Sub

Sub DeleteGradeSheetTable()
    Dim db As DAO.Database, t As DAO.TableDef
    Set db = CurrentDb
    For Each t In db.TableDefs
        If t.Name = TableName2 Then
            db.TableDefs.Delete TableName2
            Exit For
        End If
    Next t
End Sub
