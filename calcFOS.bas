Attribute VB_Name = "Macro11"
'��������� ������ � Tools/References... SolidWorks Simulation 2018 type library
Dim swApp As Object
Dim COSMOSWORKS As Object
Dim CWObject As CosmosWorksLib.CwAddincallback
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim errCode As Long '��� �������
Dim CWResult As CosmosWorksLib.cwResults
'Dim LBCMgr As CosmosWorksLib.CWLoadsAndRestraintsManager
'Dim lr As CosmosWorksLib.CWLoadsAndRestraints
Dim Face As SldWorks.Face2

'��������� ���������� ������ ������ ������ �� ������� ������
'S1_2 - ������� ���������� 1 ���� 2 (����������� ������������), ���
'S1_1 - ������� ���������� 1 ���� 1 (�������� ������������)
Public Function FOS(S1_2, S1_1, S2_2, S2_1, S3_2, S3_1 As Double)
    sn = 207 '000000 '������� �����������
    m = 1 '����������
    
    Sm3 = (S3_2 + S3_1) / 2
    Sa3 = (S3_2 - S3_1) / 2
    
    Sm2 = (S2_2 + S2_1) / 2
    Sa2 = (S2_2 - S2_1) / 2
    
    Sm1 = (S1_2 + S1_1) / 2
    Sa1 = (S1_2 - S1_1) / 2
    
    FOS = (sn - m * (Sm1 + Sm2 + Sm3) / 3) / Sqr(((Sa1 - Sa2) ^ 2 + (Sa2 - Sa3) ^ 2 + (Sa3 - Sa1) ^ 2) / 2)
End Function

Sub main()
Dim s1(1 To 2), s2(1 To 2), s3(1 To 2) As Double '������ ����������
Set swApp = Application.SldWorks '��'��� Solidworks
Set CWObject = swApp.GetAddInObject("SldWorks.Simulation") '��'��� Simulation
Set COSMOSWORKS = CWObject.COSMOSWORKS
Set ActDoc = COSMOSWORKS.ActiveDoc() '�������� �������� COSMOSWORKS
Set StudyMngr = ActDoc.StudyManager() '�������� �����
Set Part = swApp.ActiveDoc '������� ������

For Each N In Array(6313, 6334, 6349, 198, 186) '�����
For i = 1 To 2
StudyMngr.ActiveStudy = i - 1
Set Study = StudyMngr.GetStudy(i - 1) '������

'Study.MeshAndRun
'errCode = Study.CreateMesh(0, 4.7, 0.25) '�������� ����
'runError = Study.RunAnalysis '�������� ������
Set CWResult = Study.results '����������

MaxStep = CWResult.GetMaximumAvailableSteps()
sn = CWResult.GetStress(0, MaxStep, Nothing, 3, errCode) '����� ��������� � ������

s1(i) = sn((N - 1) * 12 + 7) '����� ������� ���������� (���)
s2(i) = sn((N - 1) * 12 + 8) '�����
s3(i) = sn((N - 1) * 12 + 9) '����
'Debug.Print i, s1(i), s2(i), s3(i)
Next i
Debug.Print N, FOS(s1(2), s1(1), s2(2), s2(1), s3(2), s3(1))
Next N

End Sub

