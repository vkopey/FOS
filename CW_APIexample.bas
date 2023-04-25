'Необхідно додати в Tools/References... SolidWorks Simulation 2020 type library
'Перед запуском виберіть дослідження і введіть його номер GetStudy(0), GetStudy(1)
'Ще один приклад https://help.solidworks.com/2020/English/api/swsimulationapi/Analyze_Part_Example_VB.htm
Dim swApp As Object
Dim COSMOSWORKS As Object
Dim CWObject As CosmosWorksLib.CwAddincallback
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim errCode As Long 'код помилки
Dim CWResult As CosmosWorksLib.cwResults
Dim LBCMgr As CosmosWorksLib.CWLoadsAndRestraintsManager
Dim lr As CosmosWorksLib.CWLoadsAndRestraints

'Розраховує коефіцієнт запасу втомної міцності за критерієм Сайнса
'S1_2 - головне напруження 1 крок 2 (максимальне навантаження), МПа
'S1_1 - головне напруження 1 крок 1 (мінімальне навантаження)
Public Function FOS(S1_2, S1_1, S2_2, S2_1, S3_2, S3_1 As Double)
    sn = 207 '000000 'границя витривалості
    m = 1 'коефіцієнт
    
    Sm3 = (S3_2 + S3_1) / 2
    Sa3 = (S3_2 - S3_1) / 2
    
    Sm2 = (S2_2 + S2_1) / 2
    Sa2 = (S2_2 - S2_1) / 2
    
    Sm1 = (S1_2 + S1_1) / 2
    Sa1 = (S1_2 - S1_1) / 2
    
    FOS = (sn - m * (Sm1 + Sm2 + Sm3) / 3) / Sqr(((Sa1 - Sa2) ^ 2 + (Sa2 - Sa3) ^ 2 + (Sa3 - Sa1) ^ 2) / 2)
End Function

' задіює потрібне навантаження i=1 або i=2
Sub loadActivate(i, Study)
Set LBCMgr = Study.LoadsAndRestraintsManager
Set lr1 = LBCMgr.GetLoadsAndRestraints(1, errCode)
Set lr2 = LBCMgr.GetLoadsAndRestraints(2, errCode)
If lr1.State = 0 Then lr1.SuppressUnSuppress 'якщо включений то виключити
If lr2.State = 0 Then lr2.SuppressUnSuppress 'якщо включений то виключити
Select Case i
  Case Is = 1
    lr1.SuppressUnSuppress 'включити
  Case Is = 2
    lr2.SuppressUnSuppress 'включити
End Select
Debug.Print lr1.Name, lr1.State
Debug.Print lr2.Name, lr2.State
End Sub

'напруження на елементі з іменем EntityName
Function EntityStress(Part, CWResult, EntityName)
Dim Entity As SldWorks.Edge
'Dim Entity As SldWorks.Face2
'назвіть потрібну вершину/ребро/поверхню іменем EntityName (меню FaceProperties...)
'https://help.solidworks.com/2020/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ipartdoc~getentitybyname.html
'виберіть swSelVERTICES, swSelEDGES, swSelFACES
Set Entity = Part.GetEntityByName(EntityName, swSelEDGES) 'отримати грань за назвою
'Entity.Select(0) ' вибрати (для перевірки)
Entities = Array(Entity) 'масив об'єктів
'напруження (von Mises stress, MPa) на вказаній поверхні
'масив (вузол1, значення1, вузол2, значення2,...)
'Не працює з 2D спрощенням! - код помилки 7
'https://help.solidworks.com/2020/English/api/swsimulationapi/SOLIDWORKS.Interop.cosworks~SOLIDWORKS.Interop.cosworks.ICWResults~GetStressForEntities3.html?verRedirect=1
EntityStress = CWResult.GetStressForEntities3(True, 9, 1, Nothing, (Entities), 3, 0, 0, False, errCode)
Debug.Print "Код помилки", errCode
End Function

Sub main()
Dim s1(1 To 2), s2(1 To 2), s3(1 To 2) As Double 'головні напруження
Set swApp = Application.SldWorks 'об'єкт Solidworks
Set CWObject = swApp.GetAddInObject("SldWorks.Simulation") 'об'єкт Simulation
Set COSMOSWORKS = CWObject.COSMOSWORKS
Set ActDoc = COSMOSWORKS.ActiveDoc() 'активний документ COSMOSWORKS
Set StudyMngr = ActDoc.StudyManager() 'менеджер задач
Set Study = StudyMngr.GetStudy(1) 'задача
Set Part = swApp.ActiveDoc 'активна деталь

For X = 20 To 80 Step 10 'цикл для зміни значення параметра
For i = 1 To 2
Debug.Print "Розмір, навантаження", X, i
Part.Parameter("D2@Sketch1").SystemValue = X / 1000
boolstatus = Part.EditRebuild3() 'перебудувати
loadActivate i, Study
errCode = Study.CreateMesh(0, 4.7, 0.25) 'створити сітку
runError = Study.RunAnalysis 'виконати задачу
Set CWResult = Study.results 'результати

'мінімальне і максимальне напруження за Мізесом (МПа)
smax = CWResult.GetMinMaxStress(9, 0, 1, Nothing, 3, errCode)
Debug.Print "Smax", smax(3)

'напруження в вузлі
sn = CWResult.GetStress(0, 1, Nothing, 3, errCode) 'масив напружень в вузлах
n = 65 'вузол
s1(i) = sn((n - 1) * 12 + 7) 'перше головне напруження (МПа)
s2(i) = sn((n - 1) * 12 + 8) 'друге
s3(i) = sn((n - 1) * 12 + 9) 'третє
Debug.Print "Головні напруження", s1(i), s2(i), s3(i)

'середнє значення на обраній вершині/ребрі/поверхні
ssum = 0
n = 0
For Each s In EntityStress(Part, CWResult, "Name2") 'n1, s1, n2, s2, n3, s3,...
     If VarType(s) = vbSingle Then
     ssum = ssum + s
     n = n + 1
     End If
Next
Debug.Print "Середнє S_Mises", ssum / n

Next i

'коефіцієнт FOS
Debug.Print "FOS", FOS(s1(2), s1(1), s2(2), s2(1), s3(2), s3(1))
Next X

End Sub
