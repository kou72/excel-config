Attribute VB_Name = "Module1"

Sub ExtractConfigInfo()
    ' �_�C�A���O��\�����A�I�������p�X���擾
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim selectedPath As String
    With fd
        .Title = "Select Path"
        .AllowMultiSelect = False
        
        If .Show = True Then
            selectedPath = .SelectedItems(1)
        End If
    End With

    ' �I�������p�X�� "path" �Ƃ������O�̃Z���ɏ�������
    ' Dim rng As Range
    ' Set rng = ThisWorkbook.Names("path").RefersToRange
    ' rng.Value = selectedPath

    ' Config�t�@�C�����i���ۂ̃t�@�C�����ɒu�������Ă��������j
    Dim fileName As String
    fileName = selectedPath
    
    ' �����Ώۂ̃L�[���[�h���X�g
    Dim searchWords As Variant
    searchWords = Array("interface FastEthernet0/1", "interface FastEthernet0/2", "interface FastEthernet0/3") ' �K�v�Ȃ�X�ɒǉ�
    
    ' �t�@�C����ǂݍ��݃��[�h�ŊJ��
    Dim fileNo As Integer
    fileNo = FreeFile
    Open fileName For Input As fileNo
    
    ' �t���O��������
    Dim hierarchyLevel As Integer
    Dim foundLine As Boolean
    Dim word As Variant
    Dim textLine As String
    
    ' �����L�[���[�h�Ń��[�v
    For Each word In searchWords
        hierarchyLevel = 0
        foundLine = False
        
        ' �t�@�C�����ŏ�����ǂݍ���
        Seek fileNo, 1
        
        ' �t�@�C����1�s���ǂݍ���
        Do Until EOF(fileNo)
            Line Input #fileNo, textLine
            
            ' �s���ړI�̕�������܂ނ��`�F�b�N
            If InStr(textLine, word) > 0 Then
                hierarchyLevel = Len(textLine) - Len(LTrim(textLine))
                foundLine = True
                MsgBox textLine
            ElseIf foundLine Then
                ' �ړI�̕����񂪌���������
                ' ���݂̍s���q�v�f���ǂ����`�F�b�N
                If Len(textLine) - Len(LTrim(textLine)) > hierarchyLevel Then
                    ' ����͎q�v�f�Ȃ̂ŁA�o��
                    Debug.Print textLine
                    MsgBox textLine
                Else
                    ' �q�v�f�͈̔͂𒴂����̂ŁA���[�v�𔲂���
                    Exit Do
                End If
            End If
        Loop
    Next word
    
    ' �t�@�C�������
    Close fileNo
End Sub
