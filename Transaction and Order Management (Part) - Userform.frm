VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderForm 
   Caption         =   "CityU Shop Order Form"
   ClientHeight    =   6909
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   9800.001
   OleObjectBlob   =   "Transaction and Order Management (Part) - Userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblTotal As Double, dblAccTotal

Private Sub Label4_Click()

End Sub

Private Sub lstProducts_Click()
    fmeQuantity.Enabled = True
End Sub
'Task 3 Scrollbar = Caption
Private Sub sbrQuantity_Change()
    lblQuantity.Caption = sbrQuantity.Value
End Sub

'Task 2,4
Private Sub btnOrder_Click()
Dim chrPrice As String
Dim inv As Integer

'Put the ordered item into lstOrdered
With lstOrdered
    inv = CLng(lstProducts.List(lstProducts.ListIndex, 2))
    'Task 2a In Stock
    If inv > 1 Then
        'Task 2aiii Enable fmeQuantity
        fmeQuantity.Enabled = True
        
        'Task 4a Add the name to first column
        .AddItem lstProducts.List(lstProducts.ListIndex, 0)
        chrPrice = CDbl(lstProducts.List(lstProducts.ListIndex, 1))
        sbrQuantity.Max = inv
        'Task 4b Add the quantity to the second column
        .List(lstOrdered.ListCount - 1, 1) = sbrQuantity.Value
        'Task 4c Add the amount to third column and with $ sign
        .List(lstOrdered.ListCount - 1, 2) = VBA.Format(chrPrice * sbrQuantity.Value, "$#,##0.0")
        dblTotal = dblTotal + .List(lstOrdered.ListCount - 1, 2)
        'Task 4d Change format on the label to show the total spending
        lblTotal.Caption = VBA.Format(dblTotal, "$#,##0.0")
        'Task 4e Reduce the "Item in Stock" value
        lstProducts.List(lstProducts.ListIndex, 2) = inv - sbrQuantity.Value
        sbrQuantity.Max = inv - sbrQuantity.Value
        lstProducts.ListIndex = -1
        'Task 4f Reset quantity
        sbrQuantity.Value = 1
        'Task 4g disable frame
        fmeQuantity.Enabled = False
        'The above code keeps if you would like to select the item before the quantity
        'windows shows up
    
    'Task 2b Out of stock
    Else
        'Task 2bi Error Message
        MsgBox "Selected product is out of stock. Please select another product.", , "CityU Shop"
        'Task 2bii Disable frame and make it visible
        fmeQuantity.Enabled = False
        fmeQuantity.Visible = True
    End If
    
End With

End Sub

'Task 5 Remove Preparation
Private Sub lstOrdered_Change()
Dim intIndex As Integer, intSelect As Integer

'Enable the btnRemove when any one of the items in lstOrdered is selected
For intIndex = 0 To lstOrdered.ListCount - 1
    If lstOrdered.Selected(intIndex) = True Then
        btnRemove.Enabled = True
        intSelect = intSelect + 1
    End If
Next intIndex

'Disable btnRemove if no items in lstOrdered is selected
If intSelect = 0 Then btnRemove.Enabled = False
End Sub

Private Sub btnRemove_Click()
Dim int_index, int_index1, int_found, inv As Integer
'Task 5 Remove

With lstOrdered
    'Repeat for each item in lstOrdered
    For int_index = lstOrdered.ListCount - 1 To 0 Step -1
        If .Selected(int_index) = True Then
            btnRemove.Enabled = True
            
            'Task 5a Increase the value of "Item in Stock"
            int_found = -1
            For int_index1 = lstProducts.ListCount - 1 To 0 Step -1
                'MsgBox lstProducts.List(int_index1, 0) & "+" & lstOrdered.List(int_index, 0) Testing
                If lstProducts.List(int_index1, 0) = lstOrdered.List(int_index, 0) Then
                    int_found = int_index1
                End If
            Next int_index1
            If int_found <> -1 Then
                inv = CLng(lstOrdered.List(int_index, 1))
                lstProducts.List(int_found, 2) = CLng(lstProducts.List(int_found, 2)) + inv
            End If
            
            'Task 5b Reduce the total price by the total price of the select item
            dblTotal = dblTotal - .List(int_index, 2)
            lblTotal.Caption = VBA.Format(dblTotal, "$#,##0.0")
            
            'Task 5c Remove the selected item from lstOrdered
            .RemoveItem (int_index)
            
        Else: btnRemove.Enabled = False
        
        End If
    Next int_index
End With

'Disable btnRemove
btnRemove.Enabled = False


End Sub
Private Sub btnProcOrder_Click()
'Task 6 Proceed
Dim intIndex1, sum As Long

'Task 6a Update the sales records to dblAccTotal @
For intIndex1 = 0 To lstOrdered.ListCount - 1
    dblAccTotal = dblAccTotal + CDbl(lstOrdered.List(intIndex1, 2))
Next intIndex1

'MsgBox dblAccTotal
    
'Task 6b Clear the orded items in lstOrdered
lstOrdered.Clear

'Task 6c Set dblTotal to 0 and update the caption of lblTotal
dblTotal = 0
lblTotal.Caption = VBA.Format(dblTotal, "$#,##0.0")

End Sub

Private Sub chkAccValue_Click()
'Task 7 Check total

'Task 7a Check total by msgbox
If chkAccValue.Value = True Then
    MsgBox "The accumulated order value is: " & VBA.Format(dblAccTotal, "$#,##0.0"), , "CityU Shop"
    
End If

'Task 7b Uncheck
chkAccValue.Value = False

End Sub

Private Sub btnQuit_Click()
Dim Response As Integer
'Task 8 Exit

'Task 8a Msgbox yesno
Response = MsgBox("Do you want to quit?", vbYesNo, "CityU Shop")

If Response = vbYes Then
    VBA.Unload OrderForm
    'MsgBox "Yes"
ElseIf Response = vbNo Then MsgBox "Error!"
End If

End Sub

Private Sub UserForm_Initialize()
'Tsak 1
fmeQuantity.Enabled = False
fmeQuantity.Visible = True
sbrQuantity.Max = 100
sbrQuantity.SmallChange = 1
'Task 2ai
sbrQuantity.Value = 1
dblTotal = 0
'Write your code before this line. Do not change the code below.
'**********************************************************************
Dim iCell As Long

'Activate this workbook
'ThisWorkbook.Activate

'Set up and populate lstProducts
VBA.Randomize
With lstProducts
    .ColumnCount = 3
    .ColumnWidths = "120;60;50"
  
    'For iCell = 0 To Product.Range("a1").CurrentRegion.Rows.Count - 1
    '    .AddItem Product.Range("a1").Offset(iCell, 0).Value
    '    .List(iCell, 1) = Product.Range("b1").Offset(iCell, 0).Value
    '    .List(iCell, 2) = Product.Range("c1").Offset(iCell, 0).Value
    'Next iCell
     
    .List = Product.Range("a1").CurrentRegion.Value
     
End With
'************************************************************************
End Sub








'Task 9: See the code window of ThisWorkbook object

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode <> 1 Then  'Unload statement is not involved
     Cancel = 1    'Prevent userform to be unloaded
     VBA.MsgBox "Click the Quit button to quit.", Buttons:=vbExclamation
End If

End Sub





