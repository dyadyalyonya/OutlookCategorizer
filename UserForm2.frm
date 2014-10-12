VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private strLast As String









Private Sub ListBox1_Click()

End Sub



Private Sub TextBox1_Change()

Dim tmr As Single
tmr = Timer


'Debug.Print TextBox1.SelStart & "\" & TextBox1.SelText & "\" & TextBox1.CurLine
Dim lSelStart As Long
lSelStart = TextBox1.SelStart

Dim strText As String
strText = TextBox1.text

Dim lStart As Long
lStart = InStrRev(strText, ";", lSelStart, vbTextCompare)
If lStart = 0 Then
lStart = InStrRev(strText, ",", lSelStart, vbTextCompare)
End If
'if lstart = 0 then we need 1 if we find comma or semicolon - then we need next position
lStart = lStart + 1

Dim lEnd As Long
lEnd = InStr(lSelStart, strText, ";", vbTextCompare)
If lEnd = 0 Then
lEnd = InStr(lSelStart, strText, ",", vbTextCompare)
End If
If lEnd = 0 Then
    lEnd = Len(strText)
Else
    lEnd = lEnd - 1
End If


Dim strCurCategory As String
strCurCategory = VBA.Mid(strText, lStart, lEnd - lStart + 1)
strCurCategory = VBA.Trim(strCurCategory)

Label1.Caption = lStart & " " & lEnd & " " & lSelStart & " " & strCurCategory



ListBox1.Clear
'ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(TextBox1.text, 1)
'ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(TextBox1.text, 2)
'ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(TextBox1.text, 3)
'ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(TextBox1.text, 4)
'ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(TextBox1.text, 5)
'ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(TextBox1.text, 6)
'ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(TextBox1.text, 7)

ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(strCurCategory, 1)
ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(strCurCategory, 2)
ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(strCurCategory, 3)
ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(strCurCategory, 4)
ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(strCurCategory, 5)
ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(strCurCategory, 6)
ListBox1.AddItem OUTLDATA_GetCategoryByFirstLetters(strCurCategory, 7)


ListBox1.ListIndex = 0

Debug.Print Timer - tmr

End Sub





Private Sub UserForm_Initialize()
TextBox1.text = "fuincoffer; fuoffadc   , fuoffexpir"

End Sub




'контрол состоит из текстбокса (ТБ) и списка (ЛБ)
'исх. состояние - список закрыт, текстбокс - какое-то описание категорий
'состояние правка - переход( пользователь вводит хотя бы одну букву в ТБ)- ЛБ открыт со списком катег по этой букве
'- переход (пользователь нажал ентер или стрелки право\лево?) - ЛБ закрыт в ТБ добавилась выбранная категория из ЛБ(если выбрана)
'- переход (пользователь нажал еск) - ЛБ закрыт в ТБ ничего не добавилось
'- переход (пользователь нажал стрелки V или ^) - ЛБ открыт - категория в ЛБ изменилась на одну
'- переход (пользователь мышью выбрал категорию в ЛБ) - ЛБ закрыт в ТБ добавилась категория из списка
' - переход (установка фокуса в ЛБ) - аналогично искейп?
' - переход (пользователь нажал таб) - ?

'ТБ содержит категории разделённые череж тсз или зпт
'метод - получить ту категорию которую править пользователь сейчас (может быть не последняя)
'метод - исправить ту категорию которую правили сейчас на выбранную из ЛБ




'имеем список категорий
'контроллер- текстбокс - реагриует на клавиатурный ввод пользователя
'     учитывает:
'     - положение курсора (чтобы понять какую категорию он правит сейчас)
'     - ввод (значит, что пользователь принял выбранную подсказку
'     - стрелки вверх\вниз (при активной подсказке) позволяет выбирать подсказку из списка подсказок
'      -искейп (закрывает активную посдказку) ? возможно не надо - в делишесе нет
'    возможно, чтобы перевести общение с другими объектами на более высокий уровень нужны такие события
'     - требуется подсказка (есть 1 и более непробельного символа между разделителем и курсором) - неполный текст категории
'     - нужна подсказка на 1 отличающаяся от текущей (вверх-вниз)
'     - подсказка больше не нужна (пользователь нажал искейп, либо нажал ентер - выбрал текущую предложенную посдказку)


'вьюшка - список подсказок
'     - реагирует на события контроллера
'     - реагирует на события модели
'
'модель
'         - список подсказок (описание каждой подсказки)
'         - текущая подсказка
'         - 3-4 алфавитно следующих подсказки


'модель
'Function Otlk_GetCategSuggestions( _
'        FirstLetters,
'End Function



'Function OUTLDATA_FindCategoryByFirstLetters _
'                    (ByVal strFirstLetters As String, _
'                     ByVal iOccurrence As Integer) As Outlook.Category
'
'Static colOutlCategories As Collection
'
''if we don't have index with data - build it
'If colOutCategories Is Nothing Then
'
'    'if there is no categories ever - nothing to do
'    If Application.GetNamespace("MAPI").Categories.Count = 0 Then
'        Exit Function
'    End If
'
'    'Dim varCategArray(Application.GetNamespace("MAPI").Categories.Count)
'    'Dim objCategory As Outlook.Category
'
'    'build String with all categories
'
'
'End If
'
'
'
'
'End Function





Sub Utilz_QuickSort(vArray As Variant, lLoBound As Long, lHiBound As Long)

  Dim vPivotElement   As Variant
  Dim tmpSwap As Variant
  Dim tmpLo  As Long
  Dim tmpHi   As Long
  
  If lHiBound <= lLoBound Then
    Err.Raise 1409151807, , "Can't sort empty array"
  End If

  tmpLo = lLoBound
  tmpHi = lHiBound
  
  vPivotElement = vArray((lLoBound + lHiBound) \ 2)

  While (tmpLo <= tmpHi)

     While (vArray(tmpLo) < vPivotElement And tmpLo < lHiBound)
        tmpLo = tmpLo + 1
     Wend

     While (vPivotElement < vArray(tmpHi) And tmpHi > lLoBound)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLo <= tmpHi) Then
        tmpSwap = vArray(tmpLo)
        vArray(tmpLo) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLo = tmpLo + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (lLoBound < tmpHi) Then Utilz_QuickSort vArray, lLoBound, tmpHi
  If (tmpLo < lHiBound) Then Utilz_QuickSort vArray, tmpLo, lHiBound

End Sub










'data source for "suggestion control" (suggctrl)
'находит категорию, которая начинается с переданных символов (если iOccurence = 1 то возвратит её)
'если iOccurence = 2 то возвратит следующую категорию после найденной
'если iOccurence = X то возвратить X-ую по порядку категорию после найденной
'если категория HT найдена то
'для iOccurence = 1 будет пустая строка
'для iOccurence = 2 то возвратит самую первую категорию
'для iOccurence = X возвратит X-1 категорию

'for example - user have this categories  in her outlook: {c, ccc,  ddd}
'       OUTLDATA_GetCategoryByFirstLetters ("a", 1) -> result: ""
'       OUTLDATA_GetCategoryByFirstLetters ("a", 2) -> result: "c"
'       OUTLDATA_GetCategoryByFirstLetters ("a", 3) -> result: "cсс"
'       OUTLDATA_GetCategoryByFirstLetters ("a", 4) -> result: "ddd"
'       OUTLDATA_GetCategoryByFirstLetters ("a", 5) -> result: ""
'       OUTLDATA_GetCategoryByFirstLetters ("c", 1) -> result: "c"
'       OUTLDATA_GetCategoryByFirstLetters ("c", 2) -> result: "ccc"


Function OUTLDATA_GetCategoryByFirstLetters _
                    (ByVal strFirstLetters As String, _
                     ByVal iOccurrence As Integer) As String
    
    
    
'index of all catgories by Letter
    'e.g. if we have array {c, ccc111, ccc222} we will build such collection:
    '1 key:c  value:c
    '2 key:сc  value:ccc111
    '3 key:сcc  value:ccc111
    '4 key:ccc1 value:ccc111
    '5 key:ccc2 value:ccc222
    '6 key:ccc11  value:ccc111
    '7 key:ccc22  value:ccc222
    '8 key:ccc111 value:ccc111
    '9 key:ccc222 value:ccc222
Static colOutlCtgGetNameByLetter As Collection

'index of all categories by sequential Number (get category name by seq number)
    'e.g. if we have array {c, cat111, cat222} we will build such collection:
    
    'element 1 key:"1"  value:c
    'element 2 key:"2"  value:ccc111
    'element 3 key:"3"  value:ccc222
Static colOutlCtgGetNameBySeqNum As Collection

'inverted index of all categories by sequential number (get seq number by category name)
    'e.g. if we have array {c, cat111, cat222} we will build such collection:
    
    'element 1 key:"c"  value:1
    'element 2 key:"ccc111"  value:2
    'element 3 key:"ccc222"  value:3
Static colOutlCtgGetSeqNumByName As Collection


'we don't support such values...
If iOccurrence < 1 Then Exit Function


'if we don't have indexes with data - build it
If colOutlCtgGetNameByLetter Is Nothing Then

    'if there is no categories ever - then nothing to do
    If Application.GetNamespace("MAPI").Categories.Count = 0 Then
        Exit Function
    End If
    
    'if there is more than 999999 categories - it's too much
    If Application.GetNamespace("MAPI").Categories.Count > 999999 Then
        Exit Function
    End If
    
    '0-based array of all categories
    Dim strSortedCategArray() As String
    'e.g. if we have 30 categs we redim array from 0 to 29
    ReDim strSortedCategArray(Application.GetNamespace("MAPI").Categories.Count - 1)
        
    'fill array with all categories (unsorted at this moment)
    Dim i As Long
    For i = 0 To Application.GetNamespace("MAPI").Categories.Count - 1
        strSortedCategArray(i) = Application.GetNamespace("MAPI").Categories(i + 1).Name
    Next i
    
    'sort array alphabetically
    Utilz_SortStringArray strSortedCategArray, LBound(strSortedCategArray), UBound(strSortedCategArray)
    
    'create collections
    If colOutlCtgGetNameByLetter Is Nothing Then Set colOutlCtgGetNameByLetter = New Collection
    If colOutlCtgGetNameBySeqNum Is Nothing Then Set colOutlCtgGetNameBySeqNum = New Collection
    If colOutlCtgGetSeqNumByName Is Nothing Then Set colOutlCtgGetSeqNumByName = New Collection
    
    'for each category in array...
    Dim j As Long
    For j = 0 To UBound(strSortedCategArray)
        
        colOutlCtgGetSeqNumByName.Add Item:=j + 1, key:=strSortedCategArray(j)
        colOutlCtgGetNameBySeqNum.Add Item:=strSortedCategArray(j), key:=CStr(j + 1)
    
        'for each symbol in category name....
        Dim k As Long
        For k = 1 To Len(strSortedCategArray(j))
            Dim strLeftSymbols As String
            strLeftSymbols = Left(strSortedCategArray(j), k)
            
            'trying to add key to collection
            'if we get error 457 (such key already exists) - we should go to another category
            'then we should try another suffix, else we go further
            'and add new LeftSymbols to index collection
            Err.Clear: On Error Resume Next

            colOutlCtgGetNameByLetter.Add Item:=strSortedCategArray(j), key:=strLeftSymbols
            If Err.Number = 457 Then
                'continue - key with such name already exists
            Else
                Err.Raise 1409301111, , "Forming letter index:" & Err.Description
            End If
            
        Next k
    Next j
    
End If


'searching....
On Error Resume Next


'try to find category by first letters
Dim strCtgByLetters As String
strCtgByLetters = colOutlCtgGetNameByLetter.Item(strFirstLetters)

'key was found
If Err.Number = 0 Then
    'we get category name in strCtgByLetters
    
'invalid procedure call or argument (key not found)
ElseIf Err.Number = 5 Then
    strCtgByLetters = ""
    
'Other error
Else
    Dim strErrDescr As String: strErrDescr = Err.Description
    Dim lErrNumber As Long: lErrNumber = Err.Number
    On Error GoTo 0
    Err.Raise 41012, "OUTLDATA_GetCategoryByFirstLetters", lErrNumber & " " & strErrDescr
End If



'if we want only 1st occurrence - then we've got the answer
If iOccurrence = 1 Then
    OUTLDATA_GetCategoryByFirstLetters = strCtgByLetters
    Exit Function
End If
    
    
'if we want not the 1st occurrence
'if we haven't find category by letters we
If strCtgByLetters = "" And iOccurrence > 1 Then
    On Error Resume Next 'clear ERR object
    'e.g. if they want 2 occurrence we return first category etc
    OUTLDATA_GetCategoryByFirstLetters = colOutlCtgGetNameBySeqNum(CStr(iOccurrence - 1))
    
    'key was found
    If Err.Number = 0 Then
        'we get category name in right variable, we can leave the function
        Exit Function
        
    'invalid procedure call or argument (key not found)
    ElseIf Err.Number = 5 Then
        OUTLDATA_GetCategoryByFirstLetters = ""
        Exit Function
        
    'Other error
    Else
        Dim strErrDescr11 As String: strErrDescr11 = Err.Description
        Dim lErrNumber11 As Long: lErrNumber11 = Err.Number
        On Error GoTo 0
        Err.Raise 41011, "OUTLDATA_GetCategoryByFirstLetters", lErrNumber11 & " " & strErrDescr11
    End If

End If


'if we have find category by letters then
If strCtgByLetters <> "" And iOccurrence > 1 Then
    On Error GoTo 0 'clear ERR object
    
    'error impossible here because we HAVE this category
    Dim iSeqNumber As Long
    iSeqNumber = colOutlCtgGetSeqNumByName(strCtgByLetters)
    
    On Error Resume Next 'clear ERR object
    'e.g. if they want 2 occurrence of letter "c" (example from the header) seqNumber = 2 we must return 3rd category
    OUTLDATA_GetCategoryByFirstLetters = colOutlCtgGetNameBySeqNum(CStr(iSeqNumber + iOccurrence - 1))
    
    'key was found
    If Err.Number = 0 Then
        'we get category name in right variable, we can leave the function
        Exit Function
        
    'invalid procedure call or argument (key not found)
    ElseIf Err.Number = 5 Then
        OUTLDATA_GetCategoryByFirstLetters = ""
        Exit Function
        
    'Other error
    Else
        Dim strErrDescr10 As String: strErrDescr10 = Err.Description
        Dim lErrNumber10 As Long: lErrNumber10 = Err.Number
        On Error GoTo 0
        Err.Raise 41010, "OUTLDATA_GetCategoryByFirstLetters", lErrNumber10 & " " & strErrDescr10
    End If
  
    
End If


'if we are here - my algorythm is wrong
On Error GoTo 0
Err.Raise 41009, "OUTLDATA_GetCategoryByFirstLetters", "41009 logical error"
 
    
End Function


