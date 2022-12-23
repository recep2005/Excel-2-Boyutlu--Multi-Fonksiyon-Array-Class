VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArray2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private arr_Data As Variant
Public Sub CvsdenYukle(ByVal inpFileName As String, _
                        ByVal inpRowSeper As String, _
                        Optional ByVal inpElementSeper As String = ",")

Dim i As Long
Dim j As Long
Dim tmpStr As String
Dim arr As Variant
Dim arr2 As Variant
Dim arrCSV As Variant

Open inpFileName For Binary As #1
tmpStr = Space$(LOF(1))
Get #1, , tmpStr
Close #1

arr = Split(tmpStr, inpRowSeper)

For i = LBound(arr) To UBound(arr)
    arr2 = Split(arr(i), inpElementSeper)
    
    If i = LBound(arr) Then ReDim arrCSV(LBound(arr) To UBound(arr), _
    LBound(arr2) To UBound(arr2))
    
    For j = LBound(arr2) To UBound(arr2)
        arrCSV(i, j) = arr2(j)
    Next j
Next i

arr_Data = arrCSV

End Sub



Public Property Get SatirBaslangic() As Variant

Call ArrayiKontrolEt("SatirBaslangic()")
SatirBaslangic = LBound(arr_Data, 1)
    
End Property
Public Property Get SatirBitis() As Variant

Call ArrayiKontrolEt("SatirBitis()")
SatirBitis = UBound(arr_Data, 1)
    
End Property
Public Property Get SatirSayisi() As Variant

SatirSayisi = UBound(arr_Data, 1) - LBound(arr_Data, 1) + 1

End Property
Public Property Get SutunBaslangic() As Variant

Call ArrayiKontrolEt("SutunBaslangic()")
SutunBaslangic = LBound(arr_Data, 2)

End Property
Public Property Get SutunBitis() As Variant

Call ArrayiKontrolEt("SutunBitis()")
SutunBitis = UBound(arr_Data, 2)

End Property
Public Property Get SutunSayisi() As Variant

Call ArrayiKontrolEt("SutunBitis()")
SutunSayisi = UBound(arr_Data, 2) - LBound(arr_Data, 2) + 1

End Property
Public Property Get Data() As Variant

Data = arr_Data
    
End Property

Public Property Let Data(ByRef Data As Variant)
   
If IsEmpty(Data) = True Or IsArray(Data) = False Then
    ReDim arr_Data(1 To 1, 1 To 1)
    arr_Data(1, 1) = Data
ElseIf IsArray(Data) = True Then
    If BoyutlariKontrolEt(Data) <> 2 Then
        MsgBox "Veri Array Formatinda Degil", _
        vbInformation, "Sayin " & Environ("UserName")
    End If
    arr_Data = Data
End If
    
End Property
Public Property Let Value(ByVal satir As Long, ByVal sutun As Long, ByVal Data As Variant)

arr_Data(satir, sutun) = Data

End Property
Public Property Get Value(ByVal satir As Long, ByVal sutun As Long) As Variant

Value = arr_Data(satir, sutun)

End Property
Public Sub AlandanYukle(ByVal rg As Range _
                    , Optional basligiKaldir As Boolean = False _
                    , Optional basligiSatirSayisi As Long = 1 _
                    , Optional SoldanBasligiKaldir As Boolean = False _
                    , Optional soldanBaslikSayisi As Long = 1)


If basligiKaldir = True Or SoldanBasligiKaldir = True Then

    Dim SatirSayisi As Long, rowOffset As Long
    SatirSayisi = IIf(basligiKaldir = True, rg.rows.count - _
    basligiSatirSayisi, rg.rows.count)
    rowOffset = IIf(basligiKaldir = True, basligiSatirSayisi, 0)
    
    Dim SutunSayisi As Long, columnOffset As Long
    SutunSayisi = IIf(SoldanBasligiKaldir = True, rg.columns.count - _
    soldanBaslikSayisi, rg.columns.count)
    columnOffset = IIf(SoldanBasligiKaldir = True, soldanBaslikSayisi, 0)
    
    Set rg = rg.Offset(rowOffset, columnOffset).Resize(SatirSayisi, SutunSayisi)
    
End If

arr_Data = rg.Value

End Sub
Public Function TekrarEdenleriKaldir(ByVal sadeceSecilenler As Boolean, _
                                    ParamArray SutunIndeksi()) As Variant

'Dictionary ile Calisir. Uyumluluk icin Dictionary arastirmasi yapiniz.

Dim VeriArray()
Dim ret()
Dim indexes()
Dim ids()
Dim i As Long
Dim r As Long
Dim c As Long
Dim seper As String
Dim dict As Object
Dim sutunIndexArr As Variant
Dim arrSon As Variant

If UBound(SutunIndeksi) = 0 And SutunIndeksi(0) = 0 Then
    ReDim sutunIndexArr(0 To UBound(arr_Data, 2) - 1)
    For i = LBound(sutunIndexArr) To UBound(sutunIndexArr)
        sutunIndexArr(i) = i + 1
    Next i
Else
    sutunIndexArr = SutunIndeksi
End If
Set dict = CreateObject("Scripting.Dictionary")

VeriArray = arr_Data: seper = ChrW(-1)
ReDim ids(LBound(sutunIndexArr) To UBound(sutunIndexArr))

For r = UBound(VeriArray) To LBound(VeriArray) Step -1
    For c = LBound(sutunIndexArr) To UBound(sutunIndexArr)
        ids(c) = VeriArray(r, sutunIndexArr(c))
    Next
    dict(Join$(ids, seper)) = r
Next
indexes = dict.items()
Quicksort1D indexes, LBound(indexes), UBound(indexes)

If sadeceSecilenler Then
    ReDim ret(LBound(VeriArray) To LBound(VeriArray) + dict.count - 1, _
    LBound(sutunIndexArr) + 1 To UBound(sutunIndexArr) + 1)
Else
    ReDim ret(LBound(VeriArray) To LBound(VeriArray) + dict.count - 1, _
    LBound(VeriArray, 2) To UBound(VeriArray, 2))
End If

For c = IIf(sadeceSecilenler = True, LBound(ret, 2), _
LBound(VeriArray, 2)) To IIf(sadeceSecilenler = True, _
UBound(ret, 2), UBound(VeriArray, 2))

    For r = LBound(ret) To UBound(ret)
        If sadeceSecilenler Then
            ret(r, c) = VeriArray(indexes(r - 1), sutunIndexArr(c - 1))
        Else
            ret(r, c) = VeriArray(indexes(r - 1), c)
        End If
    Next
Next

arr_Data = ret

End Function
Private Sub Quicksort1D(vArray As Variant, arrLbound As Long, arrUbound As Long)

Dim pivotVal As Variant
Dim vSwap As Variant
Dim tmpLow As Long
Dim tmpHi As Long

On Error Resume Next

tmpLow = arrLbound: tmpHi = arrUbound
pivotVal = vArray((arrLbound + arrUbound) \ 2)

While (tmpLow <= tmpHi)
    While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)
        tmpLow = tmpLow + 1
    Wend
    
    While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)
        tmpHi = tmpHi - 1
    Wend
    
    If (tmpLow <= tmpHi) Then
        vSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = vSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Wend

If (arrLbound < tmpHi) Then Quicksort vArray, arrLbound, tmpHi
If (tmpLow < arrUbound) Then Quicksort vArray, tmpLow, arrUbound
  
End Sub
Public Sub ArraydenYukle(ByVal arr As Variant)

arr_Data = arr

End Sub
Public Sub ArraydenKismiYukle(ByVal arr As Variant, _
                              Optional satir As Long = 1, _
                              Optional sutun As Long = 1, _
                              Optional SatirSayisi As Long = 1, _
                              Optional SutunSayisi As Long = 1)
                              
Dim i As Long
Dim j As Long
Dim counterRow As Long
Dim counterColumn As Long
Dim boyut As Long

boyut = BoyutlariKontrolEt(arr)
If boyut > 2 Then
    MsgBox "Array icin Sutun Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If sutun > UBound(arr, 2) Then
    MsgBox "Array icin Sutun Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If satir < LBound(arr, 1) Then
    MsgBox "Array icin Satir Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If satir + SatirSayisi > UBound(arr, 1) Then
    MsgBox "Array icin Satir Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirSayisi, _
SutunBaslangic To SutunSayisi)

counterRow = 1: counterColumn = 1
For i = LBound(arr, 2) To UBound(arr, 2)
    If i >= sutun And i + SutunSayisi - 1 Then
        For j = LBound(arr, 1) To UBound(arr, 1)
            If j >= satir And j + SatirSayisi - 1 Then
                temp(counterRow, counterColumn) = arr(j, i)
                If counterRow = UBound(temp, 1) Then counterRow = 1: Exit For
                counterRow = counterRow + 1
            End If
        Next j
        If counterColumn = UBound(temp, 2) Then Exit For
        counterColumn = counterColumn + 1
    End If
Next i

arr_Data = temp

End Sub
Public Function SutunlariSatirlardanAl(ByVal satir As Long, _
                                        Optional ByVal SatirSayisi As Long = 1) As Variant
    
If satir - 1 + SatirSayisi > SatirBitis Then
    MsgBox "Array icin Satir Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim newColumns As Variant
ReDim newColumns(1 To SutunBitis, 1 To SatirSayisi)

Dim lastRow As Long: lastRow = ((satir - 1) + SatirSayisi)
Dim currentRow As Long: currentRow = SatirBaslangic

Dim i As Long, j As Long
For i = satir To lastRow
    For j = 1 To SutunBitis
        newColumns(j, currentRow) = arr_Data(i, j)
    Next j
    currentRow = currentRow + 1
Next i

SutunlariSatirlardanAl = newColumns

End Function
Public Function SatirlariSutundanAl(ByVal sutun As Long, _
                                    Optional ByVal SutunSayisi As Long = 1) As Variant

Dim i As Long
Dim j As Long

If sutun - 1 + SutunSayisi > SutunBitis Then
    MsgBox "Array icin Sutun Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim newRows As Variant
ReDim newRows(SutunBaslangic To SutunSayisi, _
SatirBaslangic To SatirBitis)

Dim lastColumn As Long: lastColumn = ((sutun - 1) + SutunSayisi)
Dim currentColumn As Long: currentColumn = SutunBaslangic


For j = sutun To lastColumn
    For i = SatirBaslangic To SatirBitis
        newRows(currentColumn, i) = arr_Data(i, j)
    Next i
    currentColumn = currentColumn + 1
Next j

SatirlariSutundanAl = newRows

End Function
Public Function SatirlariAl(ByVal satir As Long, _
                            Optional ByVal SatirSayisi As Long = 1) As Variant


Dim newRow As Variant
ReDim newRow(SatirBaslangic To SatirSayisi, _
SutunBaslangic To SutunBitis)

Dim lastRow As Long: lastRow = ((satir - 1) + SatirSayisi)
Dim currentRow As Long: currentRow = SatirBaslangic

Dim i As Long, j As Long
For i = satir To lastRow
    For j = SutunBaslangic To SutunBitis
        newRow(currentRow, j) = arr_Data(i, j)
    Next j
    currentRow = currentRow + 1
Next i

SatirlariAl = newRow

End Function
Public Function SutunlariAl(ByVal sutun As Long, _
                            Optional ByVal SutunSayisi As Long = 1) As Variant

If sutun - 1 + SutunSayisi > SutunBitis Then
    MsgBox "Array icin Sutun Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim newColumn As Variant
ReDim newColumn(SatirBaslangic To SatirBitis, SutunBaslangic To SutunSayisi)

Dim lastColumn As Long: lastColumn = ((sutun - 1) + SutunSayisi)
Dim currentColumn As Long: currentColumn = SutunBaslangic

Dim i As Long, j As Long
For j = sutun To lastColumn
    For i = SatirBaslangic To SatirBitis
        newColumn(i, currentColumn) = arr_Data(i, j)
    Next i
    currentColumn = currentColumn + 1
Next j

SutunlariAl = newColumn

End Function
Public Sub VeriyiAlanaYazdir(ByVal rg As Range, _
                            Optional ByVal transpose As Boolean = False)
    
If transpose = True Then
    rg.Resize(SutunBitis, SatirBitis).Value = _
    SatirlariSutundanAl(1, SutunSayisi)
Else
    rg.Resize(SatirBitis, SutunBitis).Value = arr_Data
End If
    
End Sub
Public Sub SatirlariAlanaYazdir(ByVal alanaYaz As Range _
                                , ByVal satir As Long _
                                , Optional SatirSayisi As Long = 1 _
                                , Optional transpose As Boolean = False)

If transpose = True Then
    alanaYaz.Resize(SutunBitis, SatirSayisi).Value = _
    SutunlariSatirlardanAl(satir, SatirSayisi)
Else
    alanaYaz.Resize(SatirSayisi, SutunBitis).Value = _
    SatirlariAl(satir, SatirSayisi)
End If
    
End Sub
Public Sub SutunlariAlanaYazdir(ByVal alanaYaz As Range _
                                , ByVal sutun As Long _
                                , Optional SutunSayisi As Long = 1 _
                                , Optional transpose As Boolean = False)
    
If transpose = True Then
    alanaYaz.Resize(SutunSayisi, SatirBitis).Value = _
    SatirlariSutundanAl(sutun, SutunSayisi)
Else
    alanaYaz.Resize(SatirBitis, SutunSayisi).Value = _
    SutunlariAl(sutun, SutunSayisi)
End If
    
End Sub
Private Function ArrayKopyala(ByRef arrSource As Variant _
                            , ByRef arrDestination As Variant)
    
Dim i As Long, j As Long
For i = LBound(arrDestination, 1) To UBound(arrDestination, 1)
    For j = LBound(arrDestination, 2) To UBound(arrDestination, 2)
        arrDestination(i, j) = arrSource(i, j)
    Next j
Next i
    
End Function
Private Sub SatirKopyala(ByRef rowData As Variant _
                            , ByVal currentRow As Long _
                            , ByVal newRow As Long)
    
    
Dim i As Long
For i = LBound(arr_Data, 2) To UBound(arr_Data, 2)
    rowData(newRow, i) = arr_Data(currentRow, i)
Next i
    
End Sub
Private Function BoyutlariKontrolEt(ByRef arr As Variant) As Long

If IsArray(arr) = False Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

On Error GoTo errPlace

Dim dimensions As Long
Dim count As Long: count = 0

Do
    count = count + 1
    dimensions = UBound(arr, count)
Loop While True

errPlace:
BoyutlariKontrolEt = count - 1

End Function
Public Sub BosSutunEkle(ByVal sutunEkle As Long, _
                Optional ByVal SutunSayisi As Long = 1)
    
If sutunEkle > SutunBitis + 1 Then
    sutunEkle = SutunBitis + 1
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis, _
SutunBaslangic To SutunBitis + SutunSayisi)

Dim i As Long, j As Long
For j = SutunBaslangic To sutunEkle - 1
    For i = SatirBaslangic To SatirBitis
        temp(i, j) = arr_Data(i, j)
    Next i
Next j
    
For j = sutunEkle To SutunBitis
    For i = SatirBaslangic To SatirBitis
        temp(i, j + SutunSayisi) = arr_Data(i, j)
    Next i
Next j

arr_Data = temp
        
End Sub
Public Sub SutunEkleme(ByVal sutunEkle As Long, ByRef yeniSutunlar As Variant)
    
If IsEmpty(yeniSutunlar) Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If BoyutlariKontrolEt(yeniSutunlar) <> 2 Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If UBound(yeniSutunlar, 1) < SatirBitis Then
    MsgBox "Yeni Array Formati Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If sutunEkle < 1 Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If sutunEkle > SutunBitis + 1 Then
    sutunEkle = SutunBitis + 1
End If

Dim newColumnCount As Long
newColumnCount = UBound(yeniSutunlar, 2)

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis, _
SutunBaslangic To SutunBitis + newColumnCount)
Dim i As Long, j As Long, currentColumn As Long

currentColumn = sutunEkle
For j = SutunBaslangic To newColumnCount
    For i = SatirBaslangic To SatirBitis
        temp(i, currentColumn) = yeniSutunlar(i, j)
    Next i
    currentColumn = currentColumn + 1
Next j

For j = SutunBaslangic To sutunEkle - 1
    For i = SatirBaslangic To SatirBitis
        temp(i, j) = arr_Data(i, j)
    Next i
Next j
    
For j = sutunEkle To SutunBitis
    For i = SatirBaslangic To SatirBitis
        temp(i, j + newColumnCount) = arr_Data(i, j)
    Next i
Next j

arr_Data = temp
        
End Sub
Public Sub BosSatirEkle(ByVal satirEkle As Long, Optional ByVal SatirSayisi As Long = 1)
    
If satirEkle > SatirBitis + 1 Then
    satirEkle = SatirBitis + 1
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis + _
SatirSayisi, SutunBaslangic To SutunBitis)

Dim i As Long, j As Long
For i = SatirBaslangic To satirEkle - 1
    For j = SutunBaslangic To SutunBitis
        temp(i, j) = arr_Data(i, j)
    Next j
Next i
    
For i = satirEkle To SatirBitis
    For j = SutunBaslangic To SutunBitis
        temp(i + SatirSayisi, j) = arr_Data(i, j)
    Next j
Next i

arr_Data = temp
        
End Sub
Public Sub SatirEkleme(ByVal satirEkle As Long, ByRef yeniSatirlar As Variant)
    
If IsEmpty(yeniSatirlar) Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If BoyutlariKontrolEt(yeniSatirlar) <> 2 Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If UBound(yeniSatirlar, 2) < SutunBitis Then
    MsgBox "Veride Yeterli Satir Yok", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If satirEkle < 1 Then
    MsgBox "Satir Sayisi Birden Buyuk Olmalidir", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If satirEkle > SatirBitis + 1 Then
    satirEkle = SatirBitis + 1
End If

Dim newRowCount As Long
newRowCount = UBound(yeniSatirlar, 1)

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis + _
newRowCount, SutunBaslangic To SutunBitis)

Dim i As Long, j As Long, currentRow As Long
currentRow = satirEkle
For i = SatirBaslangic To newRowCount
    For j = SutunBaslangic To SutunBitis
        temp(currentRow, j) = yeniSatirlar(i, j)
    Next j
    currentRow = currentRow + 1
Next i

For i = SatirBaslangic To satirEkle - 1
    For j = SutunBaslangic To SutunBitis
        temp(i, j) = arr_Data(i, j)
    Next j
Next i
    
For i = satirEkle To SatirBitis
    For j = SutunBaslangic To SutunBitis
        temp(i + newRowCount, j) = arr_Data(i, j)
    Next j
Next i

arr_Data = temp
        
End Sub
Public Sub SutunlariKaldir(ByVal sutunuKaldir As Long, _
                            Optional ByVal SutunSayisi As Long = 1)
    
If sutunuKaldir < 1 Or sutunuKaldir > SutunBitis Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If sutunuKaldir = 1 And SutunSayisi >= SutunBitis Then
    MsgBox "Veri Sutun Sayisi Yeterli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If sutunuKaldir - 1 + SutunSayisi > SutunBitis Then
    SutunSayisi = SutunBitis - sutunuKaldir + 1
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis, _
SutunBaslangic To SutunBitis - SutunSayisi)

Dim i As Long, j As Long
For j = SutunBaslangic To sutunuKaldir - 1
    For i = SatirBaslangic To SatirBitis
        temp(i, j) = arr_Data(i, j)
    Next i
Next j

For j = sutunuKaldir + SutunSayisi To SutunBitis
    For i = SatirBaslangic To SatirBitis
        temp(i, j - SutunSayisi) = arr_Data(i, j)
    Next i
Next j

arr_Data = temp

End Sub
Public Sub SatirlariSil(ByVal satiriKaldir As Long, _
                        Optional ByVal SatirSayisi As Long = 1)
    
If satiriKaldir < 1 Or satiriKaldir > SatirBitis Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If satiriKaldir - 1 + SatirSayisi > SatirBitis Then
    MsgBox "Veri Satir Sayisi Yeterli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis - _
SatirSayisi, SutunBaslangic To SutunBitis)

Dim i As Long, j As Long
For i = SatirBaslangic To satiriKaldir - 1
    For j = SutunBaslangic To SutunBitis
        temp(i, j) = arr_Data(i, j)
    Next j
Next i

For i = satiriKaldir + SatirSayisi To SatirBitis
    For j = SutunBaslangic To SutunBitis
        temp(i - SatirSayisi, j) = arr_Data(i, j)
    Next j
Next i

arr_Data = temp

End Sub
Public Function ElementIndeks(ByVal elementAra As Variant _
                        , Optional ByVal sutun As Long = 1 _
                        , Optional ByVal baslangicPozisyonu As Long = 1 _
                        , Optional ByVal karsilastirmaTuru As VbCompareMethod = vbTextCompare _
                        , Optional ByVal bosluklariKirp As Boolean = False) As Long
    
If sutun > SutunBitis Or sutun < 1 Then
    MsgBox "Arama Elementi Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If baslangicPozisyonu > SatirBitis Then
    MsgBox "Baslangic Pozisyonu Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

ElementIndeks = -1

Dim i As Long
For i = baslangicPozisyonu To SatirBitis
    
    If StrComp(IIf(bosluklariKirp = True, Trim(arr_Data(i, sutun)), _
    arr_Data(i, sutun)), elementAra, karsilastirmaTuru) = 0 Then
        ElementIndeks = i
        Exit For
    End If
    
Next i
        
End Function
Public Function ElementSonIndeks(ByVal elementAra As Variant _
                        , Optional ByVal sutun As Long = 1 _
                        , Optional ByVal baslangicPozisyonu As Long = 1 _
                        , Optional ByVal karsilastirmaTuru As VbCompareMethod = vbTextCompare _
                        , Optional ByVal bosluklariKirp As Boolean = False) As Long
    
If sutun > SutunBitis Or sutun < 1 Then
    MsgBox "Arama Elementi Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If baslangicPozisyonu > SatirBitis Then
    MsgBox "Baslangic Pozisyonu Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

ElementSonIndeks = -1

Dim i As Long
For i = SatirBitis To baslangicPozisyonu Step -1
    
    If StrComp(IIf(bosluklariKirp = True, Trim(arr_Data(i, sutun)), _
    arr_Data(i, sutun)), elementAra, karsilastirmaTuru) = 0 Then
        ElementSonIndeks = i
        Exit For
    End If
    
Next i
        
End Function
Public Function Birlestir2DArray(IlkArray As Variant, IkinciArray As Variant) As Variant

Dim i As Long
Dim j As Long
Dim ikinciBoyut As Long
Dim arrBirlestir As Variant

ReDim arrBirlestir(1 To UBound(IlkArray, 1) + _
UBound(IkinciArray, 1), 1 To IIf(UBound(IlkArray, 2), _
UBound(IlkArray, 2), UBound(IkinciArray, 2)))

For j = LBound(IlkArray, 2) To UBound(IlkArray, 2)
    For i = LBound(IlkArray, 1) To UBound(IlkArray, 1)
        arrBirlestir(i, j) = IlkArray(i, j)
    Next i
Next j
i = 0: j = 0

ikinciBoyut = IIf(UBound(IlkArray, 2) > UBound(IkinciArray, 2), _
UBound(IlkArray, 2), UBound(IkinciArray, 2))
For j = LBound(IlkArray, 2) To ikinciBoyut
    For i = LBound(IkinciArray, 1) To UBound(IkinciArray, 1)
        arrBirlestir(i + UBound(IlkArray, 1), j) = IlkArray(i, j)
    Next i
Next j
Birlestir2DArray = arrBirlestir

End Function
Public Function Birlestir2DArrayYana(IlkArray As Variant, IkinciArray As Variant) As Variant

Dim i As Long
Dim j As Long
Dim arrBirlestir As Variant

If UBound(IlkArray, 1) > UBound(IkinciArray, 1) Then
    ReDim arrBirlestir(1 To UBound(IlkArray, 1), _
    1 To UBound(IlkArray, 2) + UBound(IkinciArray, 2))
Else
    ReDim arrBirlestir(1 To UBound(IkinciArray, 1), _
    1 To UBound(IlkArray, 2) + UBound(IkinciArray, 2))
End If

For i = LBound(IlkArray, 2) To UBound(IlkArray, 2)
    For j = LBound(IlkArray, 1) To UBound(IlkArray, 1)
        arrBirlestir(j, i) = IlkArray(j, i)
    Next j
Next i
i = 0: j = 0

For i = LBound(IkinciArray, 2) To UBound(IkinciArray, 2)
    For j = LBound(IkinciArray, 1) To UBound(IkinciArray, 1)
        arrBirlestir(j, i + UBound(IlkArray, 2)) = IkinciArray(j, i)
    Next j
Next i
Birlestir2DArrayYana = arrBirlestir

End Function
Public Function Dilimle(sadeceIlkEsleme As Boolean, ParamArray Kriter() As Variant) As Variant

Dim i As Long
Dim j As Long
Dim q As Long
Dim x As Long
Dim z As Long
Dim counter As Long
Dim matchVar As Boolean
Dim tmpVar As Variant
Dim arrTemp As Variant
Dim arr As Variant
Dim arrSonuc As Variant
Dim arrStorage As Variant
Dim arrSon As Variant
Dim arrSlicer As Variant

arr = arr_Data
arrSlicer = Kriter(0)

ReDim arrSonuc(LBound(arr, 1) To UBound(arr, 1), _
LBound(arr, 2) To UBound(arr, 2)): counter = 1
ReDim arrStorage(LBound(arr, 1) To UBound(arr, 1))

For i = LBound(arrSlicer, 1) To UBound(arrSlicer, 1)
    matchVar = True
    For j = LBound(arr, 1) To UBound(arr, 1)
        For q = LBound(Kriter) + 1 To UBound(Kriter)
            If CStr(arrSlicer(i, Split(Kriter(q), _
            "=")(UBound(Split(Kriter(q), "="))))) = _
            CStr(arr(j, Split(Kriter(q), _
            "=")(LBound(Split(Kriter(q), "="))))) Then
                matchVar = True
            Else
                matchVar = False: Exit For
            End If
        Next q
        If matchVar Then
            tmpVar = vbNullString
            For x = LBound(Kriter) + 1 To UBound(Kriter)
                tmpVar = tmpVar & ChrW(-1) & arr(j, _
                Split(Kriter(x), "=")(LBound(Split(Kriter(x), _
                "="))))
            Next x
            tmpVar = Mid(tmpVar, Len(ChrW(-1)) + 1)
            arrTemp = Split(tmpVar, ChrW(-1))
            If sadeceIlkEsleme Then
                If IsError(Application.Match(CStr(tmpVar), _
                arrStorage, 0)) Then
                     arrStorage(counter) = tmpVar
                     For z = LBound(arr, 2) To UBound(arr, 2)
                        arrSonuc(counter, z) = arr(j, z)
                     Next z
                     counter = counter + 1
                End If
            Else
                For z = LBound(arr, 2) To UBound(arr, 2)
                   arrSonuc(counter, z) = arr(j, z)
                Next z
                counter = counter + 1
            End If
        End If
    Next j
Next i

If counter = 1 Then Exit Function
ReDim arrSon(LBound(arr, 1) To counter - 1, _
LBound(arr, 2) To UBound(arr, 2)): i = 0: j = 0

For i = LBound(arrSon, 1) To UBound(arrSon, 1)
    For j = LBound(arrSon, 2) To UBound(arrSon, 2)
        arrSon(i, j) = arrSonuc(i, j)
    Next j
Next i
Dilimle = arrSon

End Function
Public Function WEBtoArray(webLink As String, seperator As String, _
                           Optional sendOptimumHeaders As Boolean = False, _
                           Optional cleanBlanks As Boolean = True, _
                           Optional cleanNonPrintingValues As Boolean = True, _
                           Optional convertTurkishHTML As Boolean = True, _
                           Optional convertTurkishtoLatin As Boolean = True, _
                           Optional toWord As Boolean = False _
                           ) As Variant
Dim i As Long
Dim responseTxt As String
Dim arr As Variant
Dim tmpVar As String

With CreateObject("MSXML2.XMLHTTP.6.0")
    .Open "GET", webLink, False
    
    If sendOptimumHeaders Then
        .setRequestHeader "Content-Type", "text/html; charset=UTF-8"
        .setRequestHeader "User-Agent", _
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " & _
        "(KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36"
    End If
    
    .Send
    Do Until .readyState = 4: DoEvents: Loop
    responseTxt = HTMLTagKaldir(.responseText)
End With

arr = Split(responseTxt, seperator): tmpVar = ""
For i = LBound(arr) To UBound(arr)
    arr(i) = IIf(cleanBlanks, Trim(arr(i)), arr(i))
    arr(i) = IIf(cleanNonPrintingValues, _
    Application.Clean(arr(i)), arr(i))
    If arr(i) <> vbNullString Then tmpVar = tmpVar & ChrW(-1) & arr(i)
Next i
Erase arr

arr = Split(Mid(tmpVar, Len(ChrW(-1)) + 1), ChrW(-1)): i = 0
For i = LBound(arr) To UBound(arr)
    If convertTurkishHTML Then
    
        arr(i) = Replace(arr(i), "&Ccedil;", Chr(199))
        arr(i) = Replace(arr(i), "&ccedil;", Chr(231))
        arr(i) = Replace(arr(i), "&#350;", Chr(222))
        arr(i) = Replace(arr(i), "&#351;", Chr(254))
        arr(i) = Replace(arr(i), "&#304;", Chr(221))
        arr(i) = Replace(arr(i), "&#305;", Chr(253))
        arr(i) = Replace(arr(i), "&Uuml;", Chr(220))
        arr(i) = Replace(arr(i), "&uuml;", Chr(252))
        arr(i) = Replace(arr(i), "&Ouml;", Chr(214))
        arr(i) = Replace(arr(i), "&ouml;", Chr(246))
        arr(i) = Replace(arr(i), "&#286;", Chr(208))
        arr(i) = Replace(arr(i), "&#287;", Chr(240))
        arr(i) = Replace(arr(i), "&nbsp", Chr(32))
        arr(i) = Replace(arr(i), "&amp;", Chr(38))
        
    End If
    If convertTurkishtoLatin Then arr(i) = TurkcetoLatin(arr(i))
Next i
If toWord Then ArraytoWord (Join(arr, vbCrLf))
WEBtoArray = arr

End Function
Private Sub ArraytoWord(inpStr As String)

Dim objWord As Object
Dim objDoc As Object
Dim objSelection

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add

With objWord
    .Visible = True
    .WindowState = 2
End With
objDoc.Activate

Set objSelection = objWord.Selection

objSelection.TypeText (inpStr)

End Sub
Private Function HTMLTagKaldir(text As String) As String

Dim regexObject As Object

With CreateObject("vbscript.regexp")
    .Pattern = "<!*[^<>]*>"
    .Global = True
    .ignoreCase = True
    .MultiLine = True
    HTMLTagKaldir = .Replace(text, "")
End With

End Function
Public Function FuzzyEslesme(ByVal aranacakKelime As Variant, _
                                ByVal aranacakSutun As Long, _
                                ByVal TurkceKaraterleriYoksay As Boolean, _
                                ByVal BosluklariYoksay As Boolean, _
                                ByVal karsilastirmaTuru As VbCompareMethod, _
                                ByVal eslesmeOrani As Double _
                                ) As Variant

Dim i As Long
Dim j As Long
Dim x As Long
Dim counter As Long
Dim matchFound As Boolean
Dim arr As Variant
Dim arrSonuc As Variant
Dim arrAktarim As Variant

arr = arr_Data: counter = 1
ReDim arrSonuc(1 To UBound(arr, 1), 1 To UBound(arr, 2))
For i = LBound(arr, 1) To UBound(arr, 1)
    For j = LBound(arr, 2) To UBound(arr, 2)
        If aranacakSutun = 0 Then GoTo skipColumnCheck
        If aranacakSutun > 0 And j = aranacakSutun Then
skipColumnCheck:
        arr(i, j) = IIf(TurkceKaraterleriYoksay, _
        TurkcetoLatin(CStr(arr(i, j))), CStr(arr(i, j)))
        arr(i, j) = IIf(BosluklariYoksay, _
        Trim(CStr(arr(i, j))), CStr(arr(i, j)))
        
        matchFound = False
        If InStr(1, CStr(arr(i, j)), CStr(aranacakKelime), _
        karsilastirmaTuru) > 0 Then
            If eslesmeOrani <= IIf(Len(aranacakKelime) > Len(arr(i, j)), _
            Len(arr(i, j)) / Len(aranacakKelime), Len(aranacakKelime) / _
            Len(arr(i, j))) Then
                matchFound = True
            End If
        End If
        If matchFound Then
            For x = LBound(arr, 2) To UBound(arr, 2)
                arrSonuc(counter, x) = arr(i, x)
            Next x
            counter = counter + 1: Exit For
        End If
        
        End If
    Next j
Next i
If counter = 1 Then Exit Function
ReDim arrAktarim(1 To counter - 1, 1 To UBound(arr, 2)): i = 0: j = 0
For i = LBound(arrAktarim, 1) To UBound(arrAktarim, 1)
    For j = LBound(arrAktarim, 2) To UBound(arrAktarim, 2)
        arrAktarim(i, j) = arrSonuc(i, j)
    Next j
Next i
FuzzyEslesme = arrAktarim

End Function
Public Function Ara(ByVal aranacakKelime As Variant, _
                    ByVal aranacakSutun As Long, _
                    ByVal TurkceKarakterleriYoksay As Boolean, _
                    ByVal BosluklariYoksay As Boolean, _
                    ByVal tamEslesme As Boolean, _
                    ByVal karsilastirmaTuru As VbCompareMethod) As Variant

Dim i As Long
Dim j As Long
Dim x As Long
Dim counter As Long
Dim matchFound As Boolean
Dim arr As Variant
Dim arrSonuc As Variant
Dim arrAktarim As Variant

arr = arr_Data: counter = 1
ReDim arrSonuc(LBound(arr, 1) To UBound(arr, 1), _
LBound(arr, 2) To UBound(arr, 2))
For i = LBound(arr, 1) To UBound(arr, 1)
    For j = LBound(arr, 2) To UBound(arr, 2)
        If aranacakSutun = 0 Then GoTo skipCheck
        If j = aranacakSutun Then
skipCheck:
                
            arr(i, j) = IIf(TurkceKarakterleriYoksay, _
            TurkcetoLatin(CStr(arr(i, j))), CStr(arr(i, j)))
            arr(i, j) = IIf(BosluklariYoksay, Trim(CStr(arr(i, j))), _
            CStr(arr(i, j)))
            
            matchFound = False
            If tamEslesme Then
                If CStr(arr(i, j)) = CStr(aranacakKelime) Then matchFound = True
            Else
                If InStr(1, CStr(arr(i, j)), CStr(aranacakKelime), _
                karsilastirmaTuru) > 0 Then matchFound = True
            End If
            
            If matchFound Then
                For x = LBound(arr, 2) To UBound(arr, 2)
                    arrSonuc(counter, x) = arr(i, x)
                Next x
                counter = counter + 1: Exit For
            End If

        End If
    Next j
Next i
If counter = 1 Then Exit Function
ReDim arrAktarim(LBound(arr, 1) To counter - 1, _
LBound(arr, 2) To UBound(arr, 2)): i = 0: j = 0

For i = LBound(arrAktarim, 1) To UBound(arrAktarim, 1)
    For j = LBound(arrAktarim, 2) To UBound(arrAktarim, 2)
        arrAktarim(i, j) = arrSonuc(i, j)
    Next j
Next i
Ara = arrAktarim

End Function
Public Function CokEFiltrele(ByVal sadeceSecilenler As Boolean, _
                            ByVal karsilastirmaTuru As VbCompareMethod, _
                            ParamArray Kriter() As Variant) As Variant

Dim i As Long
Dim j As Long
Dim q As Long
Dim counter As Long
Dim upperLimit As Long
Dim matchFound As Boolean
Dim numCompare As Boolean
Dim dateCompare As Boolean
Dim varCompare As Variant
Dim tmpVar As Variant
Dim arr As Variant
Dim arrCevap As Variant
Dim arrSon As Variant

arr = arr_Data

If sadeceSecilenler Then
    ReDim arrCevap(1 To UBound(arr, 1), 1 To UBound(Kriter) / 2)
Else
    ReDim arrCevap(1 To UBound(arr, 1), 1 To UBound(arr, 2))
End If

counter = 1
For j = LBound(arr, 1) To UBound(arr, 1)
    matchFound = True
    For i = LBound(Kriter) To UBound(Kriter)
        If i Mod 2 = 0 Then
            If Kriter(i + 1) Like ">*" Or Kriter(i + 1) Like "<*" Then
                tmpVar = Replace(Replace(Kriter(i + 1), ">", ""), "<", "")
                If IsDate(tmpVar) Then dateCompare = True Else: dateCompare = False
                If isNumeric(tmpVar) Then numCompare = True Else: numCompare = False
                If dateCompare Then varCompare = CDate(tmpVar)
                If numCompare And Not dateCompare Then varCompare = CDbl(tmpVar)
                
                If LCase(TypeName(arr(j, Kriter(i)))) = "string" And Not isNumeric(arr(j, _
                Kriter(i))) Then matchFound = False: Exit For
                arr(j, Kriter(i)) = IIf(dateCompare, CDate(arr(j, Kriter(i))), CDbl(arr(j, Kriter(i))))
                If Kriter(i + 1) Like ">*" Then
                    If Not arr(j, Kriter(i)) > varCompare Then matchFound = False: Exit For
                ElseIf Kriter(i + 1) Like "<*" Then
                    If Not arr(j, Kriter(i)) < varCompare Then matchFound = False: Exit For
                End If
            Else
                If LCase(TypeName(arr(j, Kriter(i)))) <> "string" Then matchFound = False: Exit For
                If InStr(1, Kriter(i + 1), Chr(42)) > 0 Then
                    tmpVar = Replace(Kriter(i + 1), Chr(42), "")
                    
                    If Right(Kriter(i + 1), 1) = Chr(42) And _
                    Left(Kriter(i + 1), 1) = Chr(42) Then
                        If InStr(1, arr(j, Kriter(i)), tmpVar, _
                        karsilastirmaTuru) = 0 Then matchFound = False: Exit For
                    
                    ElseIf Right(Kriter(i + 1), 1) = Chr(42) Then
                        If karsilastirmaTuru = vbTextCompare Then
                            If LCase(Left(arr(j, Kriter(i)), Len(tmpVar))) <> _
                            LCase(tmpVar) Then matchFound = False: Exit For
                        ElseIf karsilastirmaTuru = vbBinaryCompare Then
                            If Left(arr(j, Kriter(i)), Len(tmpVar)) <> tmpVar _
                            Then matchFound = False: Exit For
                        End If
                        
                    ElseIf Left(Kriter(i + 1), 1) = Chr(42) Then
                        If karsilastirmaTuru = vbTextCompare Then
                            If LCase(Right(arr(j, Kriter(i)), Len(tmpVar))) <> _
                            LCase(tmpVar) Then matchFound = False: Exit For
                        ElseIf karsilastirmaTuru = vbBinaryCompare Then
                            If Right(arr(j, Kriter(i)), Len(tmpVar)) <> _
                            tmpVar Then matchFound = False: Exit For
                        End If
                        
                    End If
                Else
                
                    If Not StrComp(arr(j, Kriter(i)), _
                    Kriter(i + 1), karsilastirmaTuru) = 0 Then
                        matchFound = False: Exit For
                    End If
                    
                End If
            End If
        End If
    Next i
    If matchFound Then
        If Not sadeceSecilenler Then
            For q = LBound(arr, 2) To UBound(arr, 2)
                arrCevap(counter, q) = arr(j, q)
            Next q
        Else
            For q = LBound(Kriter) To UBound(Kriter)
                If q Mod 2 = 1 Then
                    arrCevap(counter, Int(q / 2) + 1) = _
                    arr(j, Kriter(q - 1))
                End If
            Next q
        End If
        counter = counter + 1
    End If
Next j

If counter = 1 Then Exit Function
upperLimit = UBound(arrCevap, 2): i = 0: j = 0
ReDim arrSon(1 To counter - 1, 1 To upperLimit)

For i = LBound(arrSon, 2) To UBound(arrSon, 2)
    For j = LBound(arrSon, 1) To UBound(arrSon, 1)
        arrSon(j, i) = arrCevap(j, i)
    Next j
Next i

CokEFiltrele = arrSon

End Function
Public Function Filtrele(ByVal elementAra As Variant _
                    , Optional ByVal sutun As Long = 1 _
                    , Optional ByVal baslangicPozisyonu As Long = 1 _
                    , Optional ByVal karsilastirmaTuru As VbCompareMethod = vbTextCompare _
                    , Optional ByVal bosluklariKirp As Boolean = False) As Variant

If sutun > SutunBitis Then
    MsgBox "Arama Elementi Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If baslangicPozisyonu > SatirBitis Then
    MsgBox "Baslangic Pozisyonu Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim filteredData As Variant
ReDim filteredData(1 To UBound(arr_Data, 1), 1 To UBound(arr_Data, 2))

Dim count As Long: count = 1
    
Dim i As Long, newRow As Long, filterEmpty As Boolean
newRow = 1
filterEmpty = True
For i = baslangicPozisyonu To SatirBitis
    If StrComp(IIf(bosluklariKirp = True, Trim(arr_Data(i, sutun)), _
    arr_Data(i, sutun)), elementAra, karsilastirmaTuru) = 0 Then
        SatirKopyala filteredData, i, newRow
        newRow = newRow + 1
        filterEmpty = False
    End If
Next i

If filterEmpty = True Then
    Filtrele = Empty
Else
    Dim finalData As Variant
    ReDim finalData(1 To newRow - 1, 1 To UBound(filteredData, 2))
    ArrayKopyala filteredData, finalData
End If

Filtrele = finalData

End Function
Public Function MaxMin(ByVal sutun As Long, MaxBul As Boolean, SayisalBul As Boolean) As Variant

If sutun > SutunBitis Or sutun < 1 Then
    MsgBox "Arama Elementi Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim lnMax As Variant
Dim lnMin As Variant
Dim i As Long
Dim j As Long
Dim arr As Variant

arr = arr_Data
If SayisalBul Then
    lnMin = findNumericorNotValues(arr, sutun, True): lnMax = 0
Else
    lnMin = findNumericorNotValues(arr, sutun, False): lnMax = ""
End If

For j = LBound(arr_Data, 2) To UBound(arr_Data, 2)
    If j = sutun Then
        For i = LBound(arr_Data, 1) To UBound(arr_Data, 1)
            If SayisalBul Then
                If Not IsEmpty(arr_Data(i, j)) Then
                    If isNumeric(arr_Data(i, j)) Then
                        lnMin = IIf(arr_Data(i, j) > lnMin, lnMin, arr_Data(i, j))
                        lnMax = IIf(arr_Data(i, j) > lnMax, arr_Data(i, j), lnMax)
                    End If
                End If
            Else
                If Not isNumeric(arr_Data(i, j)) Then
                    lnMin = IIf(arr_Data(i, j) > lnMin, lnMin, arr_Data(i, j))
                    lnMax = IIf(arr_Data(i, j) > lnMax, arr_Data(i, j), lnMax)
                End If
            End If
        Next i
        Exit For
    End If
Next j

If MaxBul Then MaxMin = lnMax Else: MaxMin = lnMin

End Function
Private Function findNumericorNotValues(arr As Variant, _
                                        sutun As Long, isNumerical As Boolean)

Dim i As Long

For i = LBound(arr, 1) To UBound(arr, 1)
    If isNumerical And isNumeric(arr(i, _
    sutun)) Then findNumericorNotValues = _
    CDbl(arr(i, sutun)): Exit For
    If Not isNumerical And Not isNumeric(arr(i, _
    sutun)) Then findNumericorNotValues = _
    CStr(arr(i, sutun)): Exit For
Next i

End Function
Public Sub TerseCevir()

Dim rows As Long, columns As Long
rows = SatirBitis
columns = SutunBitis

Dim reversedData As Variant
ReDim reversedData(1 To rows, 1 To columns)

Dim i As Long, j As Long, row As Long
row = 1
For i = rows To 1 Step -1
    For j = 1 To columns
        reversedData(row, j) = arr_Data(i, j)
    Next j
    row = row + 1
Next i

arr_Data = reversedData

End Sub
Public Function CokETopla(ToplanacakSutun As Long, _
                        ParamArray Kriter() As Variant) As Variant

Dim genelToplam As Double
Dim r As Long
Dim c As Long
Dim conditionsMet As Boolean
Dim arr As Variant

arr = Me.Data: genelToplam = 0
If ToplanacakSutun < LBound(arr, 2) Or _
ToplanacakSutun > UBound(arr, 2) Then Err.Raise vbError + 9

For r = LBound(arr, 1) To UBound(arr, 1)
    conditionsMet = True
    For c = LBound(Kriter) To UBound(Kriter) - 1
        If c Mod 2 = 0 Then
            If arr(r, Kriter(c)) <> Kriter(c + 1) Then conditionsMet = False: Exit For
        End If
    Next c
    If conditionsMet And isNumeric(arr(r, ToplanacakSutun)) Then genelToplam = _
    genelToplam + arr(r, ToplanacakSutun)
Next r
CokETopla = genelToplam

End Function
Public Sub HarfBoyut(BuyukHarf As Boolean, _
                     ParamArray Kriter() As Variant)

Dim i As Long
Dim j As Long
Dim arr As Variant

arr = arr_Data
For i = LBound(arr, 1) To UBound(arr, 1)
    For j = LBound(Kriter) To UBound(Kriter)
        If LCase(TypeName(arr(i, Kriter(j)))) = "string" Then
            arr(i, Kriter(j)) = IIf(BuyukHarf, _
            UCase(arr(i, Kriter(j))), _
            LCase(arr(i, Kriter(j))))
        End If
    Next j
Next i
arr_Data = arr

End Sub
Public Sub SutunDegistir(EsasSutun As Long, DegistirilecekSutun As Long)

Dim i As Long
Dim tmpVar As Variant
Dim arr As Variant

arr = arr_Data
For i = LBound(arr, 1) To UBound(arr, 1)
    tmpVar = arr(i, DegistirilecekSutun)
    arr(i, DegistirilecekSutun) = arr(i, EsasSutun)
    arr(i, EsasSutun) = tmpVar
Next i
arr_Data = arr

End Sub
Public Sub BosluklariDoldur(kelime As Variant, ParamArray Kriter() As Variant)

Dim i As Long
Dim j As Long
Dim arr As Variant

arr = arr_Data
For j = LBound(Kriter) To UBound(Kriter)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, Kriter(j)) = vbNullString Then arr(i, Kriter(j)) = kelime
    Next i
Next j
arr_Data = arr

End Sub
Public Function CokESay(ParamArray Kriter() As Variant) As Long

Dim genelToplam As Double
Dim r As Long
Dim c As Long
Dim conditionsMet As Boolean
Dim arr As Variant

arr = Me.Data: genelToplam = 0
For r = LBound(arr, 1) To UBound(arr, 1)
    conditionsMet = True
    For c = LBound(Kriter) To UBound(Kriter) - 1
        If c Mod 2 = 0 Then
            If arr(r, Kriter(c)) <> Kriter(c + 1) Then conditionsMet = False: Exit For
        End If
    Next c
    If conditionsMet Then genelToplam = genelToplam + 1
Next r
CokESay = genelToplam

End Function
Public Function Topla(ToplanacakSutun As Long) As Variant

Dim i As Long
Dim j As Long
Dim genelToplam As Double
Dim arr As Variant

arr = arr_Data: genelToplam = 0
For j = LBound(arr, 2) To UBound(arr, 2)
    If j = ToplanacakSutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If isNumeric(arr(i, j)) Then
                genelToplam = genelToplam + arr(i, j)
            End If
        Next i
    End If
Next j
Topla = CDbl(genelToplam)

End Function
Public Function Say(SayilacakSutun As Long, _
                    Optional bosluklarHaric As Boolean = False) As Variant

Dim i As Long
Dim j As Long
Dim genelToplam As Double
Dim arr As Variant

arr = arr_Data: genelToplam = 0
For j = LBound(arr, 2) To UBound(arr, 2)
    If j = SayilacakSutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If bosluklarHaric Then
                If arr(i, j) <> "" Then
                    genelToplam = genelToplam + 1
                End If
            Else
                genelToplam = genelToplam + 1
            End If
        Next i
    End If
Next j
Say = genelToplam

End Function
Public Sub Kirp(ByVal sutun As Long, Optional SayilarDahil As Boolean = False)

Dim i As Long
Dim j As Long
Dim arr As Variant

arr = arr_Data
For j = 1 To UBound(arr, 2)
    If j = sutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If isNumeric(arr(i, j)) And Not SayilarDahil Then GoTo skipTrim
            arr(i, j) = Trim(arr(i, j))
skipTrim:
        Next i
    End If
Next j
arr_Data = arr

End Sub
Public Sub CFonksiyonlari(ByVal cTur As String, ParamArray Kriter() As Variant)

'n for numeric
's for string
'd for date

Dim i As Long
Dim j As Long
Dim arr As Variant

If cTur <> "n" And cTur <> "s" And cTur <> "d" Then Exit Sub

arr = arr_Data
For j = LBound(Kriter) To UBound(Kriter)
    For i = LBound(arr, 1) To UBound(arr, 1)
        arr(i, Kriter(j)) = IIf(cTur = "s", CStr(arr(i, Kriter(j))), _
        IIf(cTur = "n", CDbl(arr(i, Kriter(j))), CDate(arr(i, Kriter(j)))))
    Next i
Next j
arr_Data = arr

End Sub
Public Sub Degistir(ByVal sutun As Long, _
                    ByVal degisecekKelime As String, _
                    ByVal yeniKelime As String _
                    , Optional ByVal karsilastirmaTuru As VbCompareMethod = vbTextCompare _
                    , Optional ByVal TurkceKarakterleriLatinCevir As Boolean = False _
                    , Optional SayilarDahil As Boolean = False)

Dim i As Long
Dim j As Long
Dim arr As Variant

arr = arr_Data
For j = 1 To UBound(arr, 2)
    If j = sutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If isNumeric(arr(i, j)) And Not SayilarDahil Then GoTo skipTrim
            arr(i, j) = Replace(IIf(TurkceKarakterleriLatinCevir, _
            TurkcetoLatin(arr(i, j)), arr(i, j)), degisecekKelime, _
            yeniKelime, , , karsilastirmaTuru)
skipTrim:
        Next i
    End If
Next j
arr_Data = arr

End Sub
Public Sub DevrikCevir()

Dim x As Long
Dim yUbound As Long
Dim y As Long
Dim xUbound As Long
Dim tempArray As Variant
Dim arr As Variant

arr = arr_Data
xUbound = UBound(arr, 2): yUbound = UBound(arr, 1)
ReDim tempArray(1 To xUbound, 1 To yUbound)

For x = 1 To xUbound
    For y = 1 To yUbound
        tempArray(x, y) = arr(y, x)
    Next y
Next x

arr_Data = tempArray
    
End Sub
Public Function TekBoyutluyaCevir(ByVal sutun As Long) As Variant
 
Dim i As Long
Dim j As Long
Dim tmpVar As String
Dim arr As Variant
Dim arr1D As Variant

arr = arr_Data
ReDim arr1D(1 To UBound(arr, 1))

For j = LBound(arr, 2) To UBound(arr, 2)
    If j = sutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            arr1D(i) = arr(i, j)
        Next i
    End If
Next j
TekBoyutluyaCevir = arr1D

End Function
Public Function IkiBoyutluyaCevir(ByVal arr As Variant) As Variant
 
Dim i As Long
Dim oneArray As Boolean
Dim arr2D As Variant
 
If LBound(arr) = 0 Then
    oneArray = False
    ReDim arr2D(1 To UBound(arr) + 1, 1 To 1)
    For i = LBound(arr2D, 1) To UBound(arr2D, 1)
        arr2D(i, 1) = arr(i - 1)
    Next i
Else
    oneArray = True
    ReDim arr2D(1 To UBound(arr), 1 To 1)
    For i = LBound(arr2D, 1) To UBound(arr2D, 1)
        arr2D(i, 1) = arr(i)
    Next i
End If
IkiBoyutluyaCevir = arr2D

End Function
Public Sub TurkceKarakterleriCevir(BosluklariKalsin As Boolean, ParamArray Kriter() As Variant)

Dim i As Long
Dim j As Long
Dim tmpVar As String
Dim arr As Variant

arr = arr_Data
For j = LBound(Kriter) To UBound(Kriter)
    If j + 1 <= UBound(arr, 2) Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If Not isNumeric(arr(i, Kriter(j))) Then
                tmpVar = IIf(BosluklariKalsin = True, _
                arr(i, Kriter(j)), Trim(arr(i, Kriter(j))))
                arr(i, Kriter(j)) = TurkcetoLatin(tmpVar)
            End If
        Next i
    End If
Next j
arr_Data = arr

End Sub
Private Function TurkcetoLatin(ByVal strInput As String) As String

Dim i As Long
Dim arrTurkce() As Variant
Dim arrEnglish() As Variant

arrTurkce = Array(Chr(253), Chr(221), Chr(240), Chr(208), Chr(252), Chr(220), _
Chr(254), Chr(222), Chr(246), Chr(214), Chr(231), Chr(199))
arrEnglish = Array("i", "I", "g", "G", "u", "U", "s", "S", "o", "O", "c", "C")

For i = LBound(arrTurkce) To UBound(arrTurkce)
    If InStr(1, strInput, arrTurkce(i), vbBinaryCompare) > 0 Then
        strInput = Replace(strInput, arrTurkce(i), arrEnglish(i))
    End If
Next i
TurkcetoLatin = strInput

End Function
Public Sub Sirala(ByVal sutun As Long, _
                  Optional ByVal harfDuyarsiz As Boolean = True)

Quicksort SortArray:=arr_Data, _
lngColumn:=sutun, harfDuyarsiz:=harfDuyarsiz

End Sub
Private Sub Quicksort(ByRef SortArray As Variant, Optional lngMin As Long = -1, _
                       Optional lngMax As Long = -1, Optional lngColumn As Long = 0, _
                       Optional harfDuyarsiz As Boolean = True)
    
On Error Resume Next

Dim i As Long
Dim j As Long
Dim varMid As Variant
Dim arrRowTemp As Variant
Dim lngColTemp As Long

If IsEmpty(SortArray) Then Exit Sub
If InStr(TypeName(SortArray), "()") < 1 Then Exit Sub
If lngMin = -1 Then lngMin = LBound(SortArray, 1)
If lngMax = -1 Then lngMax = UBound(SortArray, 1)
If lngMin >= lngMax Then Exit Sub

i = lngMin: j = lngMax: varMid = Empty
varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

If IsObject(varMid) Then
    i = lngMax
    j = lngMin
ElseIf IsEmpty(varMid) Then
    i = lngMax
    j = lngMin
ElseIf IsNull(varMid) Then
    i = lngMax
    j = lngMin
ElseIf varMid = "" Then
    i = lngMax
    j = lngMin
ElseIf VarType(varMid) = vbError Then
    i = lngMax
    j = lngMin
ElseIf VarType(varMid) > 17 Then
    i = lngMax
    j = lngMin
End If

While i <= j
    If harfDuyarsiz Then
        While UCase(SortArray(i, lngColumn)) < UCase(varMid) And i < lngMax
            i = i + 1
        Wend
        While UCase(varMid) < UCase(SortArray(j, lngColumn)) And j > lngMin
            j = j - 1
        Wend
        
    Else
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend
    End If
    
    If i <= j Then
        ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
        For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
            arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
            SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
            SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
        Next lngColTemp
        Erase arrRowTemp

        i = i + 1
        j = j - 1
    End If
Wend

If (lngMin < j) Then Call Quicksort(SortArray, lngMin, j, lngColumn, harfDuyarsiz)
If (i < lngMax) Then Call Quicksort(SortArray, i, lngMax, lngColumn, harfDuyarsiz)

End Sub
Public Sub ArrayiKontrolEt(ByVal procedure As String)

If IsEmpty(arr_Data) Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

End Sub
