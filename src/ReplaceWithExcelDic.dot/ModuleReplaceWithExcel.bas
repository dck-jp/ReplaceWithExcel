Attribute VB_Name = "ModuleReplaceWithExcel"
Option Explicit
' <original code>
'   ���[�h�����̌�Q��Excel�̎������g���Ēu�� ver.04
' Author: �n�Ӑ^
' URL: http://makoto-watanabe.main.jp/WordVba_replace.html
'
' modified By D*isuke YAMAKAWA
' last modified: 2012/8/1

Const xlUp = -4162

'Excel �����t�@�C�� �֘A�̕ϐ� _________________
Const colNumBefore As Integer = 1 '�u���Ώۂ̌����L�ڂ����
Const colNumAfter As Integer = 2 '�u����̌����L�ڂ����
Const colNumEnablesWildcards As Integer = 3 '���C���h�J�[�h�g�p�̗L�����L�ڂ����
Const IsTrue As String = "�L��" '���C���h�J�[�h�̎g�p������������

Const rowNumBeginData As Integer = 2 '�����f�[�^�̋L�ڂ��n�܂�s
' _______________________________________________


Sub Excel�̎������g���Ēu��()
    Dim filename As String
    Dim dic() As ReplaceData '�����f�[�^���i�[����z��
    Dim sw As Stopwatch
    Set sw = New Stopwatch
    
    '�����t�@�C����I��
    filename = SelectDictionary
    If filename = vbNullString Then Exit Sub
    
    sw.Start '�������Ԃ̌v���J�n
        dic = LoadDictionary(filename) '�z��Ɏ����f�[�^��ǂݍ���
        Call SortByWordCount(dic) '�����f�[�^����בւ�
        Call ReplaceWithDictionary(dic) '�����f�[�^���g���Ēu��
    sw.Stop_ '�������Ԃ̌v���I��

    MsgBox "�������I�����܂����B" & vbNewLine & "�������Ԃ́A" _
        & Format(sw.Elapsed, "hh����nn��ss�b") & " �ł����B"
End Sub

'�����f�[�^��p���āAWord�t�@�C���̒u�����s��
'@dictionary() : �����f�[�^
Private Sub ReplaceWithDictionary(dictionary() As ReplaceData)
    Application.ScreenUpdating = False
    
    Dim i
    For i = 0 To UBound(dictionary)
        '�u����̌�傪�A�i����͂Łj�󔒂ɂȂ��Ă��Ȃ����m�F
        If dictionary(i).afterWord <> "" Then
            With Selection.Find
               .ClearFormatting
               .text = dictionary(i).beforeWord
               .Forward = True
               .Wrap = wdFindContinue
               .Format = False
               .MatchCase = False
               .MatchByte = False
               .MatchSoundsLike = False
               .MatchAllWordForms = False
               .MatchFuzzy = False
               
               If dictionary(i).EnablesWildcard Then
                  .MatchWildcards = True
               Else
                  .MatchWildcards = False
               End If
            
               .Replacement.ClearFormatting
               .Replacement.text = dictionary(i).afterWord
               .Font.Hidden = False
            End With
            
            Selection.Find.Execute Replace:=wdReplaceAll
            
            'Word�̉�ʂ��ł܂�̂�h������
            '�P�̌��̒u�����I��邽�тɃ��b�Z�[�W�|���v�������I�ɉ�
            DoEvents
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub

'�����f�[�^���ꐔ���i�~���j�ɕ��ёւ���
'�\�[�g�ɗp�����A���S���Y��: Gnome Sort
'@jisyo() : �����f�[�^
Public Sub SortByWordCount(ByRef jisyo() As ReplaceData)
    Dim temp As ReplaceData
    Dim i As Integer, n As Integer
    
    i = 1
    n = UBound(jisyo) + 1
    
    Do While i < n
        If (jisyo(i - 1).wordCount < jisyo(i).wordCount) Then
            temp = jisyo(i - 1)
            jisyo(i - 1) = jisyo(i)
            jisyo(i) = temp
            i = i - 1
            
            If i = 0 Then
                i = 2
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

'�����t�@�C����z��Ɋi�[���āA�擾����
'@ filename : Excel �����t�@�C����
Private Function LoadDictionary(filename As String) As ReplaceData()
    'Excel�̎����t�@�C�����J��
    Dim excelApp, workBook
    Set excelApp = CreateObject("Excel.Application")
    Set workBook = excelApp.Workbooks.Open(filename)
    
        '�����t�@�C���Ƀf�[�^���L�ڂ���Ă���ŏI�s���擾
        Dim rowNumLastData As Integer
        rowNumLastData = excelApp.cells(workBook.ActiveSheet.Rows.count, 1).End(xlUp).Row
        
        '�擾�����ŏI�s�Ɋ�Â��Ď����t�@�C���̃f�[�^�������肵�A
        '�����f�[�^���i�[����z����m�ۂ���
        Dim dictionaryData() As ReplaceData
        Dim masterDataCount
        masterDataCount = rowNumLastData - rowNumBeginData + 1
        ReDim dictionaryData(masterDataCount)
        
        '�����t�@�C������s�Âz��Ɋi�[����
        Dim i, currentRowNum
        For i = 0 To masterDataCount
            currentRowNum = i + rowNumBeginData
        
            Dim repData As ReplaceData
            Dim temp
            repData.afterWord = excelApp.cells(currentRowNum, colNumAfter).Value
            repData.beforeWord = excelApp.cells(currentRowNum, colNumBefore).Value
            
            temp = excelApp.cells(currentRowNum, colNumEnablesWildcards).Value
            If temp = IsTrue Then
                repData.EnablesWildcard = True
            Else
                repData.EnablesWildcard = False
            End If
            
            repData.wordCount = Len(repData.beforeWord)
            
            If repData.afterWord <> "" Then
                dictionaryData(i) = repData
            End If
        Next i
    
    '�����t�@�C�������
    workBook.Close SaveChanges:=False
    excelApp.Quit
    Set excelApp = Nothing
    
    LoadDictionary = dictionaryData
End Function

'�t�@�C���I���_�C�A���O���J���āA�I�����ꂽ�t�@�C�������擾����
'@return : �t�@�C�����i�t���p�X�j
'          ���t�@�C����I�����Ȃ������ꍇvbNullString��Ԃ�
Private Function SelectDictionary() As String
    MsgBox "Excel �����t�@�C����I�����Ă��������B"
    
    Dim dlg As Dialog
    Dim dlgFind As Dialog
    Set dlg = Dialogs(wdDialogFileOpen)
    Set dlgFind = Dialogs(wdDialogFileFind)

    With dlg
        .Name = "*.xls"
        Select Case .Display
        Case -1 '�t�@�C�����I�����ꂽ�Ƃ�
            dlgFind.Update
            SelectDictionary = dlgFind.SearchPath & "\" & dlg.Name
        Case Else '�L�����Z���{�^���������ꂽ�Ƃ�
            SelectDictionary = vbNullString
        End Select
    
    End With
End Function

