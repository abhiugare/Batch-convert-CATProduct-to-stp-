
Private Sub CommandButton1_Click()
With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Input Folder"
        If .Show = -1 Then ' If the user selects a folder
            TextBox1.Value = .SelectedItems(1) ' Store the path in TextBox1
        End If
    End With
End Sub

Private Sub CommandButton2_Click()
With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Output Folder"
        If .Show = -1 Then ' If the user selects a folder
            TextBox2.Value = .SelectedItems(1) ' Store the path in TextBox2
        End If
    End With
End Sub

Private Sub CommandButton3_Click()
   ' Call the main process function with paths from TextBoxes
    Call StartProcessing(TextBox1.Value, TextBox2.Value)
    Me.Hide ' Hide the UserForm after processing
End Sub


Sub CATMain()

    ' Show the UserForm to get the user input paths
        Dim UserForm As New UserForm1
        UserForm.Show
    
     ' When the user presses the ProcessButton, the processing function is called
     ' The StartProcessing function will use the paths from the UserForm

End Sub

Sub StartProcessing(inputFolderPath As String, outputFolderPath As String)

    Dim cat As Application
    Set cat = CATIA.Application
    
    Dim inputpath As String
    Dim outputpath As String
    
    inputpath = inputFolderPath
    outputpath = outputFolderPath
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fs.GetFolder(inputpath)
    
    Dim i As Integer
    i = 1
    
    
    Dim subfolder As Object
    Dim file As Object
    
        For Each subfolder In folder.SubFolders
        
            Dim pdoc As ProductDocument
            Set pdoc = cat.Documents.Add("Product")
        
            Dim product1 As Product
            Set product1 = pdoc.Product
        
            product1.PartNumber = "Product_" & i
        
            For Each file In subfolder.Files
        
                If LCase(fs.GetExtensionName(file.Name)) = "catproduct" Then
                         
                     Dim products1 As Products
                     Set products1 = product1.Products
            
                     Dim arrayOfVariantOfBSTR1(0)
                     arrayOfVariantOfBSTR1(0) = file.Path
                     
                     Set products1Variant = products1
                     products1Variant.AddComponentsFromFiles arrayOfVariantOfBSTR1, "All"
                    
                End If
                
            Next
                    
            Dim outfilepath As String
            outfilepath = outputpath & "\" & "Product_" & i & ".step"
            pdoc.ExportData outfilepath, "stp"
            pdoc.Close
            MsgBox (i & " file converted")
            i = i + 1
            
        Next

End Sub


