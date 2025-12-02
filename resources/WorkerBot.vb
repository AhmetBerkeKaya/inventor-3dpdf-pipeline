' Bu kural Processor.ipt dosyasına 'WorkerBot' adıyla eklenmelidir.
' Tetikleyici: "After Open Document"

Sub Main()
    ' CONFIG: Yolların Python ile eşleştiğinden emin olun
    Dim jobFile As String = "C:\3DPDF_Pipeline\Temp\job.txt"
    Dim logFile As String = "C:\3DPDF_Pipeline\Temp\worker_log.txt"
    Dim stlFile As String = "C:\3DPDF_Pipeline\Temp\temp_export.stl"

    ' İş yoksa çık
    If Not System.IO.File.Exists(jobFile) Then Exit Sub
    
    Dim fileToProcess As String = System.IO.File.ReadAllText(jobFile).Trim()
    If fileToProcess = "" Then Exit Sub
    
    ' Dosyayı sil ve logla
    System.IO.File.Delete(jobFile)
    System.IO.File.AppendAllText(logFile, "Starting Job: " & fileToProcess & vbCrLf)
    
    ' Gizli parça oluştur
    Dim oDoc As PartDocument = ThisApplication.Documents.Add(kPartDocumentObject, , False)
    
    Try
        Dim ext As String = System.IO.Path.GetExtension(fileToProcess).ToLower()
        
        ' --- STRATEJİLER ---
        If ext = ".rvt" Then
            ' Revit: ImportedComponent (AnyCAD)
            Dim oDef As ImportedComponentDefinition
            oDef = oDoc.ComponentDefinition.ReferenceComponents.ImportedComponents.CreateDefinition(fileToProcess)
            oDoc.ComponentDefinition.ReferenceComponents.ImportedComponents.Add(oDef)
            
        ElseIf ext = ".dwg" Or ext = ".dxf" Then
            ' DWG: ACAD Translator (Generic Object)
            ' Bu yöntem 2D/3D ayrımı yapmadan dosyayı açar.
            Dim oAddIns As ApplicationAddIns = ThisApplication.ApplicationAddIns
            Dim oAcadTrans As TranslatorAddIn = oAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
            
            If Not oAcadTrans.Activated Then oAcadTrans.Activate()
            
            Dim oTransContext As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext
            oTransContext.Type = kDataDropIOMechanism
            Dim oTransOptions As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap
            Dim oTransData As DataMedium = ThisApplication.TransientObjects.CreateDataMedium
            oTransData.FileName = fileToProcess
            
            ' Dokümanı çeviriciye açtırıyoruz
            Dim oResultObject As Object = Nothing
            oAcadTrans.Open(oTransData, oTransContext, oTransOptions, oResultObject)
            
            ' Eğer açılan şey Parça değilse (Örn: Teknik Resim), STL olamaz.
            If Not oResultObject Is Nothing Then
                If TypeOf oResultObject Is PartDocument Then
                    ' Parça ise eski doc'u kapat, yenisine geç
                    oDoc.Close(True) 
                    oDoc = oResultObject 
                Else
                    Throw New Exception("WARNING: File opened as Drawing (2D). Skipping.")
                End If
            End If
            
        Else
            ' IPT/IAM: Derived Part
            Dim oDef As DerivedPartDefinition
            oDef = oDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(fileToProcess)
            oDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(oDef)
        End If
        
        ' STL EXPORT
        Dim oStlTrans As TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")
        If Not oStlTrans.Activated Then oStlTrans.Activate()
        
        Dim oContextSTL As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext()
        oContextSTL.Type = kFileBrowseIOMechanism
        Dim oOptionsSTL As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
        Dim oDataMediumSTL As DataMedium = ThisApplication.TransientObjects.CreateDataMedium()
        oDataMediumSTL.FileName = stlFile
        
        If oStlTrans.HasSaveCopyAsOptions(oDoc, oContextSTL, oOptionsSTL) Then
            oOptionsSTL.Value("ExportUnits") = 2
            oOptionsSTL.Value("Resolution") = 0
            oOptionsSTL.Value("OutputFileType") = 0
        End If
        
        If System.IO.File.Exists(stlFile) Then System.IO.File.Delete(stlFile)
        
        oStlTrans.SaveCopyAs(oDoc, oContextSTL, oOptionsSTL, oDataMediumSTL)
        System.IO.File.AppendAllText(logFile, "STL Export Success." & vbCrLf)
        
    Catch ex As Exception
        System.IO.File.AppendAllText(logFile, "ERROR: " & ex.Message & vbCrLf)
    Finally
        If Not oDoc Is Nothing Then oDoc.Close(True)
    End Try
End Sub