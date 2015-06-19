VERSION 5.00
Begin VB.Form FilterByFormDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter By Form"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox allowDebugChk 
      Caption         =   "Get Query Statement"
      Height          =   195
      Left            =   5520
      TabIndex        =   33
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton btnPrintScreen 
      Caption         =   "Print Screen"
      Height          =   555
      Left            =   7470
      TabIndex        =   32
      Top             =   6090
      Width           =   855
   End
   Begin VB.ComboBox AndOr 
      Height          =   315
      Index           =   6
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   5040
      Width           =   1245
   End
   Begin VB.ComboBox FieldSelect 
      Height          =   315
      Index           =   7
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   30
      Top             =   5400
      Width           =   3825
   End
   Begin VB.TextBox TextToMatch 
      Height          =   285
      Index           =   7
      Left            =   4020
      TabIndex        =   29
      Top             =   5400
      Width           =   2625
   End
   Begin VB.ComboBox AndOr 
      Height          =   315
      Index           =   5
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   4290
      Width           =   1245
   End
   Begin VB.ComboBox FieldSelect 
      Height          =   315
      Index           =   6
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   4650
      Width           =   3825
   End
   Begin VB.TextBox TextToMatch 
      Height          =   285
      Index           =   6
      Left            =   4020
      TabIndex        =   26
      Top             =   4650
      Width           =   2625
   End
   Begin VB.TextBox TextToMatch 
      Height          =   285
      Index           =   5
      Left            =   4020
      TabIndex        =   24
      Top             =   3900
      Width           =   2625
   End
   Begin VB.ComboBox FieldSelect 
      Height          =   315
      Index           =   5
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   3900
      Width           =   3825
   End
   Begin VB.ComboBox AndOr 
      Height          =   315
      Index           =   4
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3540
      Width           =   1245
   End
   Begin VB.TextBox TextToMatch 
      Height          =   285
      Index           =   4
      Left            =   4020
      TabIndex        =   21
      Top             =   3150
      Width           =   2625
   End
   Begin VB.ComboBox FieldSelect 
      Height          =   315
      Index           =   4
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   3150
      Width           =   3825
   End
   Begin VB.ComboBox AndOr 
      Height          =   315
      Index           =   3
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2790
      Width           =   1245
   End
   Begin VB.ComboBox FieldSelect 
      Height          =   315
      Index           =   3
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   3825
   End
   Begin VB.TextBox TextToMatch 
      Height          =   285
      Index           =   3
      Left            =   4020
      TabIndex        =   10
      Top             =   2400
      Width           =   2625
   End
   Begin VB.ComboBox AndOr 
      Height          =   315
      Index           =   2
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2040
      Width           =   1245
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Cancel"
      Height          =   525
      Left            =   3870
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1485
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   525
      Left            =   1590
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1425
   End
   Begin VB.TextBox hiddenOwner 
      Height          =   285
      Left            =   60
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "hidden calling form name"
      Top             =   6300
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TextToMatch 
      Height          =   285
      Index           =   2
      Left            =   4020
      TabIndex        =   7
      Top             =   1650
      Width           =   2625
   End
   Begin VB.TextBox TextToMatch 
      Height          =   285
      Index           =   1
      Left            =   4020
      TabIndex        =   4
      Top             =   900
      Width           =   2625
   End
   Begin VB.TextBox TextToMatch 
      Height          =   285
      Index           =   0
      Left            =   4020
      TabIndex        =   1
      Top             =   150
      Width           =   2625
   End
   Begin VB.ComboBox AndOr 
      Height          =   315
      Index           =   1
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1290
      Width           =   1245
   End
   Begin VB.ComboBox AndOr 
      Height          =   315
      Index           =   0
      ItemData        =   "FilterByFormDialog.frx":0000
      Left            =   450
      List            =   "FilterByFormDialog.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   1245
   End
   Begin VB.ComboBox FieldSelect 
      Height          =   315
      Index           =   2
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1650
      Width           =   3825
   End
   Begin VB.ComboBox FieldSelect 
      Height          =   315
      Index           =   1
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   900
      Width           =   3825
   End
   Begin VB.ComboBox FieldSelect 
      Height          =   315
      Index           =   0
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   3825
   End
   Begin VB.Label Label6 
      Caption         =   "for date fields, don't use = comparisons, only <= or >="
      Height          =   585
      Left            =   6720
      TabIndex        =   25
      Top             =   4770
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Sales Rank: must be ""dollars"" or ""units"", rank A/B/C/D"
      Height          =   585
      Left            =   6750
      TabIndex        =   18
      Top             =   4110
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "for qty*(specific whse) need to give whse # then comparison+qty i.e. ""000 <=7"""
      Height          =   855
      Left            =   6750
      TabIndex        =   17
      Top             =   3180
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "checkboxes should be ""yes"" or ""no"""
      Height          =   405
      Left            =   6750
      TabIndex        =   16
      Top             =   2700
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "text fields can use NOT, BEGINSWITH, ENDSWITH, EXACTLY, IS EMPTY, IS NOT EMPTY (must be all caps), otherwise anywhere in string"
      Height          =   1635
      Left            =   6750
      TabIndex        =   15
      Top             =   990
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "numeric fields can use <>, <, >, <=, >=, if nothing then it assumes ="
      Height          =   825
      Left            =   6750
      TabIndex        =   14
      Top             =   90
      Width           =   1515
   End
End
Attribute VB_Name = "FilterByFormDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ADDING A NEW SET
'  - copy + paste new control
'  - update size of FBF_* arrays in GlobalsPoinv.bas
'  - no error catching, so make sure it's right, otherwise it'll reference a nonexistent
'    array cell or control.

Private whereClause As String
Private havingClause As String
Private cType As String
Private i As Long
Private fromOK As Boolean

Private Sub btnExit_Click()
    Unload FilterByFormDialog
End Sub

Private Sub btnOK_Click()
    'save choices
    For i = 0 To UBound(FBF_FIELDS)
        FBF_FIELDS(i) = Me.FieldSelect(i)
        FBF_TEXT(i) = Me.TextToMatch(i)
    Next i
    For i = 0 To UBound(FBF_ANDOR)
        FBF_ANDOR(i) = Me.AndOr(i)
    Next i

    Dim strsql As String
    whereClause = ""
    havingClause = ""
    For i = 0 To Me.FieldSelect.count - 1
        If Me.FieldSelect(i) = "" Or Me.TextToMatch(i) = "" Then
            i = Me.FieldSelect.count 'bail
        Else
            handleClauses
        End If
    Next i
    
    
    strsql = "SELECT DISTINCT InventoryMaster.ItemNumber FROM InventoryMaster "
    
    If InStr(whereClause, "ProductTag") Then
        strsql = strsql & " INNER JOIN ProductTags ON ProductTags.ItemNumber = InventoryMaster.ItemNumber "
    End If
    
    

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Tom's Code That Added Filter To Sort By "Last Inventoried Date"
   ' Added 7/8/2014
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If InStr(whereClause, "DynamicPricing.ItemNumber LIKE '%t%'") > 0 Then
        strsql = strsql & ", DynamicPricing "
       whereClause = Replace(whereClause, "DynamicPricing.ItemNumber LIKE '%t%'", "DynamicPricing.ItemNumber = InventoryMaster.ItemNumber")
    End If
    If InStr(whereClause, "DynamicPricing.Enabled LIKE '%t%'") > 0 Then
        strsql = strsql & ", DynamicPricing "
        whereClause = Replace(whereClause, "DynamicPricing.Enabled LIKE '%t%'", "DynamicPricing.Enabled=1 AND DynamicPricing.ItemNumber = InventoryMaster.ItemNumber")
        
    
    End If
    
    If InStr(whereClause, "LastInventoriedDate") Then
        strsql = strsql & "INNER JOIN InventoryComponentMap ON InventoryComponentMap.ItemNumber = InventoryMaster.ItemNumber "
        strsql = strsql & " INNER JOIN LocationContents ON LocationContents.ComponentID = InventoryComponentMap.ComponentID "
       
    End If
    
    'Tom's Code 9-24-2014 Adding Subtitle Filter
    If InStr(whereClause, "Subtitle Overridden Item (t/f)") Then
        
        
            whereClause = Replace(whereClause, "Subtitle Overridden Item (t/f)", "")
            'whereClause = whereClause & " InventoryMaster.EBayPrice>200"
            If InStr(whereClause, "SubtitledItems") = False Then
                'if whereClause ends in 'AND' then remove it...
                Dim clauseLen As Integer
                clauseLen = Len(whereClause)
                Dim clauseClause As String
                If Len(whereClause) > 3 Then
                clauseClause = Mid(whereClause, clauseLen - 3, 3)
                If clauseClause = "AND" Then
                    whereClause = Mid(whereClause, 1, clauseLen - 4)
                End If
                If Mid(whereClause, 1, 4) = " AND" Then
                    whereClause = Mid(whereClause, 5, clauseLen)
                End If
                If InStr(whereClause, " AND  AND") > 0 Then
                    whereClause = Replace(whereClause, " AND  AND", " AND")
                End If
                End If
                strsql = strsql & "INNER JOIN SubtitledItems ON SubtitledItems.ItemNumber=InventoryMaster.ItemNumber "
            End If
             'whereClause = "InventoryMaster.ItemNumber not in(Select ItemNumber FROM SubtitledItems)"
        'Else
            'whereClause = "InventoryMaster.ItemNumber NOT IN (SELECT ItemNumber FROM SubtitledItems)"
        'End If
        
    End If
    
    If InStr(whereClause, "Subtitled By Default (true)") > 0 Or InStr(whereClause, "Subtitled By Default (false)") > 0 Then
        Dim whereClauseEnd As String
        Dim whereClauseBeginning As String
        
        whereClauseEnd = Mid(whereClause, Len(whereClause) - 2, 3)
        
        If whereClauseEnd <> "AND" Then whereClause = whereClause & " AND "
        
        If InStr(whereClause, "(true)") > 0 Then
            whereClause = Replace(whereClause, "Subtitled By Default (true)", "")
            strsql = strsql & "INNER JOIN vInventorySpreadsheet ON vInventorySpreadsheet.ITEMNUMBER=InventoryMaster.ItemNumber "
             whereClause = whereClause & "(vInventorySpreadsheet.EBayPrice > 200 OR TotalSalesLast365Days>50) AND InventoryMaster.ItemNumber NOT IN (SELECT ItemNumber FROM SubtitledItems) "
        End If
        If InStr(whereClause, "(false)") > 0 Then
            whereClause = Replace(whereClause, "Subtitled By Default (false)", "")
            strsql = strsql & "INNER JOIN vInventorySpreadsheet ON vInventorySpreadsheet.ITEMNUMBER=InventoryMaster.ItemNumber "
            whereClause = whereClause & "vInventorySpreadsheet.EBayPrice < 201 AND TotalSalesLast365Days<51 "
        End If
        
        whereClause = Trim(whereClause)
        whereClauseBeginning = Mid(whereClause, 1, 3)
        If whereClauseBeginning = "AND" Then whereClause = Mid(whereClause, 4, Len(whereClause))
        'whereClause = Replace(whereClause, "Subtitled By Default (t/f)", "")
        'strsql = strsql & "INNER JOIN PartNumbers ON PartNumbers.ItemNumber = InventoryMaster.ItemNumber "
        'strsql = strsql & "INNER JOIN vInventorySpreadsheet ON vInventorySpreadsheet.ITEMNUMBER=InventoryMaster.ItemNumber "
        'whereClause = whereClause & "vInventorySpreadsheet.EBayPrice > 200 OR TotalSalesLast365Days>50"
    End If
   
   If InStr(whereClause, "Subtitled (true)") > 0 Or InStr(whereClause, "Subtitled (false)") > 0 Then
        Dim whereClauseEnd2 As String
        Dim whereClauseBeginning2 As String
        whereClauseBeginning2 = Mid(whereClause, 1, 3)
        If whereClauseBeginning2 = "AND" Then whereClause = Mid(whereClause, 4, Len(whereClause))
        whereClauseEnd2 = Mid(whereClause, Len(whereClause) - 2, 3)
        If whereClauseEnd2 <> "AND" Then whereClause = whereClause & " AND "
        
        If InStr(whereClause, "Subtitled (true)") > 0 Then
        whereClause = Replace(whereClause, "Subtitled (true)", "")
        strsql = strsql & "INNER JOIN vInventorySpreadsheet ON vInventorySpreadsheet.ITEMNUMBER=InventoryMaster.ItemNumber "
        strsql = strsql & "INNER JOIN SubtitledItems ON SubtitledItems.ItemNumber=InventoryMaster.ItemNumber "
        whereClause = whereClause & "(vInventorySpreadsheet.EBayPrice > 200 OR TotalSalesLast365Days>50) OR InventoryMaster.ItemNumber IN (SELECT ItemNumber FROM SubtitledItems)"
        End If
        
        If InStr(whereClause, "Subtitled (false)") > 0 Then
        whereClause = Replace(whereClause, "Subtitled (false)", "")
        strsql = strsql & "INNER JOIN vInventorySpreadsheet ON vInventorySpreadsheet.ITEMNUMBER=InventoryMaster.ItemNumber "
        whereClause = whereClause & "vInventorySpreadsheet.EBayPrice < 201 AND TotalSalesLast365Days<51 AND InventoryMaster.ItemNumber NOT IN (SELECT ItemNumber FROM SubtitledItems)"
        End If
                
        whereClause = Trim(whereClause)
        
        If Mid(whereClause, 1, 3) = "AND" Then
            whereClause = Mid(whereClause, 4, Len(whereClause))
        End If

        whereClause = Replace(whereClause, "AND  AND", "AND")
   
   
   End If
    
    If InStr(whereClause, "WarehouseID") Then
    
    
        If InStr(strsql, "InventoryComponentMap ON") < 1 Then
        strsql = strsql & " INNER JOIN InventoryComponentMap ON InventoryMaster.ItemNumber = InventoryComponentMap.ItemNumber"
        End If
        If InStr(strsql, "LocationContents ON") < 1 Then
        strsql = strsql & " INNER JOIN LocationContents ON InventoryComponentMap.ComponentID = LocationContents.ComponentID"
        End If
        If InStr(strsql, "LocationMaster ON") < 1 Then
        strsql = strsql & " INNER JOIN LocationMaster ON LocationMaster.ID=LocationContents.LocationID"
        End If
        
        If InStr(strsql, "InventoryComponents ON") < 1 Then
        strsql = strsql & " INNER JOIN InventoryComponents ON InventoryComponentMap.ComponentID=InventoryComponents.ID"
        End If
        

        
        
        
        
        
    End If
    
    
    If InStr(whereClause, "SphereOneCategories.Name") Then
        strsql = strsql & " INNER JOIN PartNumberWebPaths ON PartNumberWebPaths.ItemNumber=InventoryMaster.ItemNumber INNER JOIN WebPaths ON WebPaths.ID = PartNumberWebPaths.WebPathID INNER JOIN MasterCategoryConnectors ON MasterCategoryConnectors.MasterID = PartNumberWebPaths.WebPathID INNER JOIN SphereOneCategories ON SphereOneCategories.ID = MasterCategoryConnectors.ConnectorID "
        whereClause = whereClause & " AND MasterCategoryConnectors.ConnectorType=4" ' AND SphereOneCategories.Name LIKE"
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    
    
    
    
    If InStr(whereClause, "PartNumbers.") Then
        strsql = strsql & " LEFT OUTER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber"
    End If
    If InStr(whereClause, "AP_Vendor.") Then
        strsql = strsql & " LEFT OUTER JOIN AP_Vendor ON InventoryMaster.PrimaryVendor=AP_Vendor.VendorNo"
    End If
    If InStr(whereClause, "SignNames.") Then
        If Not CBool(InStr(strsql, "PartNumbers")) Then
            strsql = strsql & " LEFT OUTER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber"
        End If
        strsql = strsql & " LEFT OUTER JOIN SignNames ON PartNumbers.SignID=SignNames.ID"
    End If
    'If InStr(whereClause, "GroupSigns.") Then
    '    strsql = strsql & " LEFT OUTER JOIN PartNumberGroupSigns ON InventoryMaster.ItemNumber=PartNumberGroupSigns.ItemNumber INNER JOIN GroupSigns ON GroupSigns.ID=PartNumberGroupSigns.GroupID"
    'End If
    If InStr(whereClause, "PartNumberSpecials.") Then
        strsql = strsql & " LEFT OUTER JOIN PartNumberSpecials ON InventoryMaster.ItemNumber=PartNumberSpecials.ItemNumber"
    End If
    If InStr(whereClause, "AvailOverrides.") Then
        If Not CBool(InStr(strsql, "PartNumbers")) Then
            strsql = strsql & " LEFT OUTER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber"
        End If
        strsql = strsql & " LEFT OUTER JOIN AvailOverrides ON PartNumbers.AvailabilityOverride=AvailOverrides.ID"
    End If
    If InStr(whereClause, "PartNumberAccessories.") Then
        strsql = strsql & " LEFT OUTER JOIN PartNumberAccessories ON InventoryMaster.ItemNumber=PartNumberAccessories.ItemNumber"
    End If
    If InStr(whereClause, "vActualWhseQty.") Then
        strsql = strsql & " INNER JOIN vActualWhseQty ON InventoryMaster.ItemNumber=vActualWhseQty.ItemNumber"
    End If
    If InStr(whereClause, "InventoryQuantities.") Or InStr(havingClause, "InventoryQuantities.") Then
        strsql = strsql & " INNER JOIN InventoryQuantities ON InventoryMaster.ItemNumber=InventoryQuantities.ItemNumber"
    End If
    If InStr(whereClause, "SalesRankings") Then
        strsql = strsql & " INNER JOIN SalesRankings ON InventoryMaster.ItemNumber=SalesRankings.ItemNumber"
    End If
    If InStr(whereClause, "vItemStatusBreakdown.") Then
        strsql = strsql & " INNER JOIN vItemStatusBreakdown ON InventoryMaster.ItemNumber=vItemStatusBreakdown.ItemNumber"
    End If
    'If InStr(whereClause, "TmpWeight.") Then
    '    'TODO: this is using inventorymaster weight, not component weight
    '    strsql = strsql & " LEFT OUTER JOIN (SELECT ItemNumber, CASE WHEN CHARINDEX(';',Weight)=0 THEN CAST(Weight AS float) ELSE -1 END AS WeightNumeric FROM InventoryMaster) AS TmpWeight ON InventoryMaster.ItemNumber=TmpWeight.ItemNumber"
    'End If
    If InStr(whereClause, "InventoryComponents.") Then
        strsql = strsql & " LEFT OUTER JOIN InventoryComponentMap ON InventoryMaster.ItemID=InventoryComponentMap.ItemID"
        strsql = strsql & " LEFT OUTER JOIN InventoryComponents ON InventoryComponentMap.ComponentID=InventoryComponents.ID"
    End If
    If InStr(whereClause, "vInventoryComponentsDisplayable.") Then
        strsql = strsql & " LEFT OUTER JOIN vInventoryComponentsDisplayable ON InventoryMaster.ItemNumber=vInventoryComponentsDisplayable.ItemNumber"
    End If
    If InStr(whereClause, "ISDAW0.") Then 'sales static
        strsql = strsql & " LEFT OUTER JOIN vInventoryStatisticsDenormalizedAllWhses AS ISDAW0 ON InventoryMaster.ItemNumber=ISDAW0.ItemNumber AND ISDAW0.PeriodType=0"
    End If
    If InStr(whereClause, "ISDAW1.") Then 'sales rolling
        strsql = strsql & " LEFT OUTER JOIN vInventoryStatisticsDenormalizedAllWhses AS ISDAW1 ON InventoryMaster.ItemNumber=ISDAW1.ItemNumber AND ISDAW1.PeriodType=1"
    End If
    If InStr(havingClause, "vBarcodes.") Then
        strsql = strsql & " LEFT OUTER JOIN vBarcodes ON InventoryMaster.ItemNumber=vBarcodes.ItemNumber"
    End If
    If InStr(whereClause, "InventoryLocationInfo.") Then
        'TODO: won't work with multiple, either figure that out or ref a denormalized view
        strsql = strsql & " LEFT OUTER JOIN InventoryLocationInfo ON InventoryMaster.ItemNumber=InventoryLocationInfo.ItemNumber"
    End If
    If InStr(havingClause, "PartNumberAccessories.") Then
        strsql = strsql & " LEFT OUTER JOIN PartNumberAccessories ON InventoryMaster.ItemNumber=PartNumberAccessories.ItemNumber"
    End If
    If InStr(whereClause, "vCurrentPurchaseOrderInfo.") Then
        strsql = strsql & " LEFT OUTER JOIN vCurrentPurchaseOrderInfo ON InventoryMaster.ItemNumber=vCurrentPurchaseOrderInfo.ItemNumber"
    End If
    If InStr(whereClause, "vInventoryLocationInfoSummed.") Then
        strsql = strsql & " LEFT OUTER JOIN vInventoryLocationInfoSummed ON InventoryMaster.ItemNumber=vInventoryLocationInfoSummed.ItemNumber"
    End If
    If InStr(whereClause, "PartNumberGroupLines.") Then
        strsql = strsql & " LEFT OUTER JOIN PartNumberGroupLines ON InventoryMaster.ItemNumber=PartNumberGroupLines.ItemNumber"
    End If
    If InStr(whereClause, "EBayCategoryE.") Then
        If Not CBool(InStr(strsql, "PartNumbers")) Then
            strsql = strsql & " LEFT OUTER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber"
        End If
        strsql = strsql & " LEFT OUTER JOIN EBayCategory AS EBayCategoryE ON PartNumbers.EBayCategoryID=EBayCategoryE.CategoryID AND EBayCategoryE.CategoryType=0"
    End If
    If InStr(whereClause, "EBayCategoryS.") Then
        If Not CBool(InStr(strsql, "PartNumbers")) Then
            strsql = strsql & " LEFT OUTER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber"
        End If
        strsql = strsql & " LEFT OUTER JOIN EBayCategory AS EBayCategoryS ON PartNumbers.EBayStoreCategoryID=EBayCategoryS.CategoryID AND EBayCategoryS.CategoryType=1"
    End If
    'more tables go here
    
    If whereClause <> "" Then
        strsql = strsql & " WHERE " & whereClause
    End If
   
    If havingClause <> "" Then
        strsql = strsql & " GROUP BY InventoryMaster.ItemNumber HAVING " & havingClause
    End If
    
    
  
       strsql = strsql & " ORDER BY InventoryMaster.ItemNumber"
       'MsgBox strsql
    
  
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Tom's code here as not to interfere with previous workings of adding a "clause".
    'Ended in epic failure... simply put...
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    'Dim datelastinv As Integer
    
'    For datelastinv = 0 To Me.FieldSelect.count - 1 Step 1

 '       If Me.FieldSelect(datelastinv) = "Date Last Inventoried" Then
            'Dim selectedDate As String
 '           selectedDate = Mid(Me.TextToMatch(datelastinv), 2, Len(Me.TextToMatch(datelastinv)))
            'Dim datelastinv_MathSign As String
            
 '           If Mid(selectedDate, 1, 1) = "=" Then
 '               datelastinv_MathSign = Mid(Me.TextToMatch(datelastinv), 1, 2)
 '               selectedDate = Right(selectedDate, Len(selectedDate) - 1)
                
 '           Else
 '               datelastinv_MathSign = Mid(Me.TextToMatch(datelastinv), 1, 1)
 '           End If
            
 '           strsql = "SELECT ItemNumber,LocationID,LocationContents.ComponentID,LocationContents.Quantity,LastInventoriedDate FROM LocationContents JOIN InventoryComponentMap ON LocationContents.ComponentID = InventoryComponentMap.ComponentID WHERE" & " DATEDIFF(day,LastInventoriedDate,'" & selectedDate & "')" & datelastinv_MathSign & "0 ORDER BY InventoryComponentMap.ItemNumber"
            
 '       End If
  '  Next datelastinv
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   
   
   
   
   
   
   
   
   
    
    If Me.hiddenOwner = "InventoryMaintenance" Then
        InventoryMaintenance.btnFilterByForm.tag = strsql
    ElseIf Me.hiddenOwner = "SignMaintenance" Then
        SignMaintenance.btnFilterByForm.tag = strsql
    End If
    
    
        
        If allowDebugChk = vbChecked Then
        DebugOutputFrm.WriteOutput strsql
        fromOK = True
         Unload FilterByFormDialog
         DebugOutputFrm.Show
         Else
         fromOK = True
         Unload FilterByFormDialog
         
        End If
   
   
    
        'InventoryMaintenance.CheckDynamicPriceEnabled
   
        
End Sub

Private Sub btnPrintScreen_Click()
    ClipBoard_SetImage Utilities.CaptureVB6Form(Me)
    MsgBox "Screenshot captured!"
End Sub

Private Sub FieldSelect_Click(index As Integer)
    FieldSelect_LostFocus index
End Sub

Private Sub FieldSelect_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.FieldSelect(index), KeyCode, Shift
End Sub

Private Sub FieldSelect_KeyPress(index As Integer, KeyAscii As Integer)
    AutoCompleteKeyPress Me.FieldSelect(index), KeyAscii
    If KeyAscii = vbKeyReturn Then
        FieldSelect_LostFocus index
    End If
End Sub

Private Sub FieldSelect_LostFocus(index As Integer)
    AutoCompleteLostFocus Me.FieldSelect(index)
End Sub

Private Sub Form_Load()
    Dim i As Long
    For i = 0 To Me.AndOr.count - 1
        Me.AndOr(i).AddItem "AND"
        Me.AndOr(i).AddItem "OR"
        Me.AndOr(i) = "AND"
    Next i
    populateCombos
    For i = 0 To UBound(FBF_FIELDS)
        Me.FieldSelect(i) = FBF_FIELDS(i)
        Me.TextToMatch(i) = FBF_TEXT(i)
    Next i
    For i = 0 To UBound(FBF_ANDOR)
        Me.AndOr(i) = IIf(FBF_ANDOR(i) = "", "AND", FBF_ANDOR(i))
    Next i
    fromOK = False
End Sub

Private Sub populateCombos()
    For i = 0 To Me.FieldSelect.count - 1
        Me.FieldSelect(i).AddItem "Item Description"
        Me.FieldSelect(i).AddItem "Product Line"
        Me.FieldSelect(i).AddItem "Vendor Name"
        Me.FieldSelect(i).AddItem "Vendor Number"
        Me.FieldSelect(i).AddItem "Additional Info"
        Me.FieldSelect(i).AddItem "Notes"
        'Me.FieldSelect(i).AddItem "Weight"
        Me.FieldSelect(i).AddItem "Weight (any)"
        Me.FieldSelect(i).AddItem "Weight (total)"
        Me.FieldSelect(i).AddItem "Length"
        Me.FieldSelect(i).AddItem "Width"
        Me.FieldSelect(i).AddItem "Height"
        'Me.FieldSelect(i).AddItem "Dimensions"
        'Me.FieldSelect(i).AddItem "Replaced By"
        'Me.FieldSelect(i).AddItem "Replacement For"
        Me.FieldSelect(i).AddItem "Replaced By (t/f)"
        Me.FieldSelect(i).AddItem "Replacement For (t/f)"
        Me.FieldSelect(i).AddItem "Shipping Class"
        Me.FieldSelect(i).AddItem "Discontinued"
        Me.FieldSelect(i).AddItem "Discontinued (web)"
        Me.FieldSelect(i).AddItem "Being Discontinued"
        Me.FieldSelect(i).AddItem "Dropship (any d/s)"
        Me.FieldSelect(i).AddItem "Dropship Only"
        Me.FieldSelect(i).AddItem "Standard Pack"
        Me.FieldSelect(i).AddItem "Standard Cost"
        Me.FieldSelect(i).AddItem "Average Cost"
        Me.FieldSelect(i).AddItem "Store Price"
        Me.FieldSelect(i).AddItem "List Price"
        Me.FieldSelect(i).AddItem "Last Price Change"
        Me.FieldSelect(i).AddItem "MAPP"
        Me.FieldSelect(i).AddItem "TNC"
        Me.FieldSelect(i).AddItem "PO Comment"
        Me.FieldSelect(i).AddItem "New Order Quantity"
        Me.FieldSelect(i).AddItem "Quantity On Hand"
        Me.FieldSelect(i).AddItem "Quantity On Sales Order"
        Me.FieldSelect(i).AddItem "Quantity On Purchase Order"
        Me.FieldSelect(i).AddItem "Quantity On Back Order"
        Me.FieldSelect(i).AddItem "Quantity On Hand (specific whse)"
        Me.FieldSelect(i).AddItem "Quantity On Sales Order (specific whse)"
        Me.FieldSelect(i).AddItem "Quantity On Purchase Order (specific whse)"
        Me.FieldSelect(i).AddItem "Quantity On Back Order (specific whse)"
        Me.FieldSelect(i).AddItem "Quantity On Hand (Vendor)"
        Me.FieldSelect(i).AddItem "Current Order Point"
        'Me.FieldSelect(i).AddItem "New Order Point"
        'Me.FieldSelect(i).AddItem "Locked Order Point"
        'Me.FieldSelect(i).AddItem "Using New Order Point"
        Me.FieldSelect(i).AddItem "Sales This Month (Total)"
        Me.FieldSelect(i).AddItem "Sales Last Month (Total)"
        Me.FieldSelect(i).AddItem "Sales This Period (Total)"
        Me.FieldSelect(i).AddItem "Sales Last Period (Total)"
        Me.FieldSelect(i).AddItem "Sales Year To Date (Total)"
        Me.FieldSelect(i).AddItem "Sales Last Year (Total)"
        Me.FieldSelect(i).AddItem "Sales This Month (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Last Month (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales This Period (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Last Period (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Year To Date (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Last Year (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales This Month (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Last Month (SO Only)"
        Me.FieldSelect(i).AddItem "Sales This Period (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Last Period (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Year To Date (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Last Year (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 30 (Total)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 30 to 60 Ago (Total)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 90 (Total)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 90 to 180 (Total)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 365 (Total)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 365 to 730 (Total)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 30 (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 30 to 60 Ago (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 90 (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 90 to 180 (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 365 (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 365 to 730 (P2 Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 30 (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 30 to 60 Ago (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 90 (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 90 to 180 (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 365 (SO Only)"
        Me.FieldSelect(i).AddItem "Sales Rolling Last 365 to 730 (SO Only)"
        Me.FieldSelect(i).AddItem "Returns This Month"
        Me.FieldSelect(i).AddItem "Returns Last Month"
        Me.FieldSelect(i).AddItem "Returns This Period"
        Me.FieldSelect(i).AddItem "Returns Last Period"
        Me.FieldSelect(i).AddItem "Returns Year To Date"
        Me.FieldSelect(i).AddItem "Returns Last Year"
        Me.FieldSelect(i).AddItem "Returns Rolling Last 30"
        Me.FieldSelect(i).AddItem "Returns Rolling Last 30 to 60"
        Me.FieldSelect(i).AddItem "Returns Rolling Last 90"
        Me.FieldSelect(i).AddItem "Returns Rolling Last 90 to 180"
        Me.FieldSelect(i).AddItem "Returns Rolling Last 365"
        Me.FieldSelect(i).AddItem "Returns Rolling Last 365 to 730"
        Me.FieldSelect(i).AddItem "In-Store Quantity Break 1"
        Me.FieldSelect(i).AddItem "In-Store Quantity Break 2"
        Me.FieldSelect(i).AddItem "In-Store Quantity Break 3"
        Me.FieldSelect(i).AddItem "In-Store Quantity Break 4"
        Me.FieldSelect(i).AddItem "In-Store Quantity Break 5"
        Me.FieldSelect(i).AddItem "In-Store Price Break 1"
        Me.FieldSelect(i).AddItem "In-Store Price Break 2"
        Me.FieldSelect(i).AddItem "In-Store Price Break 3"
        Me.FieldSelect(i).AddItem "In-Store Price Break 4"
        Me.FieldSelect(i).AddItem "In-Store Price Break 5"
        Me.FieldSelect(i).AddItem "Internet Quantity Break 1"
        Me.FieldSelect(i).AddItem "Internet Quantity Break 2"
        Me.FieldSelect(i).AddItem "Internet Quantity Break 3"
        Me.FieldSelect(i).AddItem "Internet Quantity Break 4"
        Me.FieldSelect(i).AddItem "Internet Quantity Break 5"
        Me.FieldSelect(i).AddItem "Internet Price Break 1"
        Me.FieldSelect(i).AddItem "Internet Price Break 2"
        Me.FieldSelect(i).AddItem "Internet Price Break 3"
        Me.FieldSelect(i).AddItem "Internet Price Break 4"
        Me.FieldSelect(i).AddItem "Internet Price Break 5"
        'Tom's Code 1-12-2015
        'to add Item Number Contains field
        Me.FieldSelect(i).AddItem "Item Number Contains"
        Me.FieldSelect(i).AddItem "Description 1"
        Me.FieldSelect(i).AddItem "Description 2"
        Me.FieldSelect(i).AddItem "Description 3"
        Me.FieldSelect(i).AddItem "Description 4"
        Me.FieldSelect(i).AddItem "Spec 1"
        Me.FieldSelect(i).AddItem "Spec 2"
        Me.FieldSelect(i).AddItem "Spec 3"
        Me.FieldSelect(i).AddItem "Spec 4"
        Me.FieldSelect(i).AddItem "Spec 5"
        Me.FieldSelect(i).AddItem "Spec 6"
        Me.FieldSelect(i).AddItem "Spec 7"
        Me.FieldSelect(i).AddItem "Spec 8"
        Me.FieldSelect(i).AddItem "Ext. Specs - Additional Info"
        Me.FieldSelect(i).AddItem "Ext. Specs - Write Up"
        Me.FieldSelect(i).AddItem "Ext. Specs - Features"
        Me.FieldSelect(i).AddItem "Ext. Specs - Specs"
        Me.FieldSelect(i).AddItem "Ext. Specs - Media"
        Me.FieldSelect(i).AddItem "Ext. Specs - Important Info"
        Me.FieldSelect(i).AddItem "Web Orderable"
        Me.FieldSelect(i).AddItem "To Be Published"
        Me.FieldSelect(i).AddItem "Published"
        Me.FieldSelect(i).AddItem "Web Delete?"
        Me.FieldSelect(i).AddItem "Web Deleted"
        Me.FieldSelect(i).AddItem "Path 1"
        Me.FieldSelect(i).AddItem "Path 2"
        Me.FieldSelect(i).AddItem "Path 3"
        Me.FieldSelect(i).AddItem "Path 4"
        Me.FieldSelect(i).AddItem "Path 5"
        Me.FieldSelect(i).AddItem "Accessory"
        Me.FieldSelect(i).AddItem "Accessory Count"
        Me.FieldSelect(i).AddItem "Availability"
        Me.FieldSelect(i).AddItem "Availability Limit"
        Me.FieldSelect(i).AddItem "Sign"
        'Me.FieldSelect(i).AddItem "Group Sign"
        Me.FieldSelect(i).AddItem "Special"
        Me.FieldSelect(i).AddItem "Section Special URL"
        Me.FieldSelect(i).AddItem "Section Special Image"
        Me.FieldSelect(i).AddItem "Print Sign"
        Me.FieldSelect(i).AddItem "Display On Front"
        Me.FieldSelect(i).AddItem "Web Caption"
        Me.FieldSelect(i).AddItem "Sales Rank (This Period)"
        Me.FieldSelect(i).AddItem "Sales Rank (Last Period)"
        Me.FieldSelect(i).AddItem "Sales Rank (Year to Date)"
        Me.FieldSelect(i).AddItem "Sales Rank (Last Year)"
        Me.FieldSelect(i).AddItem "Triad Code"
        Me.FieldSelect(i).AddItem "Truck Freight"
        Me.FieldSelect(i).AddItem "Truck Freight $6.50"
        Me.FieldSelect(i).AddItem "Reconditioned"
        Me.FieldSelect(i).AddItem "Needs A Box"
        Me.FieldSelect(i).AddItem "Request Form Only"
        Me.FieldSelect(i).AddItem "Pricing Critical (Non-Show)"
        Me.FieldSelect(i).AddItem "Pricing Critical (Show)"
        Me.FieldSelect(i).AddItem "Altered Item/To Be Offloaded"
        Me.FieldSelect(i).AddItem "Barcode (y/n)"
        Me.FieldSelect(i).AddItem "Status - Stock"
        Me.FieldSelect(i).AddItem "Status - Dropship"
        Me.FieldSelect(i).AddItem "Status - D/C by Us"
        Me.FieldSelect(i).AddItem "Status - D/C by Mfgr"
        Me.FieldSelect(i).AddItem "GMROI"
        Me.FieldSelect(i).AddItem "Web/Phone Price"
        Me.FieldSelect(i).AddItem "Stock - Store"
        Me.FieldSelect(i).AddItem "Stock - Whse"
        Me.FieldSelect(i).AddItem "OP - Store"
        Me.FieldSelect(i).AddItem "OP - Whse"
        Me.FieldSelect(i).AddItem "Follies (t/f)"
        Me.FieldSelect(i).AddItem "Below OP (t/f)"
        Me.FieldSelect(i).AddItem "Meta No-Index"
        Me.FieldSelect(i).AddItem "Date Published"
        Me.FieldSelect(i).AddItem "Date Reviewed"
        Me.FieldSelect(i).AddItem "Date Entered"
        
        
        
        'Tom's Code below to add the Date Last Inventoried Field in dropdown lists
        Me.FieldSelect(i).AddItem "Date Last Inventoried"
        Me.FieldSelect(i).AddItem "Warehouse ID"
        Me.FieldSelect(i).AddItem "Subtitle Overridden Item (t/f)" 'shows all items in SubtitledItems
        Me.FieldSelect(i).AddItem "Subtitled By Default (t/f)" 'shows all items EBay Price > 200 OR SalesRollingLast365Days(total) > 50
        Me.FieldSelect(i).AddItem "Subtitled (t/f)" ' shows all items that have a subtitle whether defaulted or not.
        Me.FieldSelect(i).AddItem "Dynamic Pricing Available (t/f)"
        Me.FieldSelect(i).AddItem "Dynamic Pricing Enabled (t/f)"
        Me.FieldSelect(i).AddItem "EBay Out Of Stock (t/f)"
        Me.FieldSelect(i).AddItem "Sphere 1 Category Name"
        
        'Tom's Code below to add Product Tag filtering...
        Me.FieldSelect(i).AddItem "Product Tag Like"
        
        Me.FieldSelect(i).AddItem "Profit Margin % (StdCost / Store)"
        Me.FieldSelect(i).AddItem "Profit Margin % (TNC / Store)"
        Me.FieldSelect(i).AddItem "Profit Margin % (AvgCost / Store)"
        Me.FieldSelect(i).AddItem "Profit Margin % (StdCost / Web)"
        Me.FieldSelect(i).AddItem "Profit Margin % (TNC / Web)"
        Me.FieldSelect(i).AddItem "Profit Margin % (AvgCost / Web)"
        Me.FieldSelect(i).AddItem "Profit Margin $ (StdCost / Store)"
        Me.FieldSelect(i).AddItem "Profit Margin $ (TNC / Store)"
        Me.FieldSelect(i).AddItem "Profit Margin $ (AvgCost / Store)"
        Me.FieldSelect(i).AddItem "Profit Margin $ (StdCost / Web)"
        Me.FieldSelect(i).AddItem "Profit Margin $ (TNC / Web)"
        Me.FieldSelect(i).AddItem "Profit Margin $ (AvgCost / Web)"
        
        Me.FieldSelect(i).AddItem "Grouped Master Item (t/f)"
        Me.FieldSelect(i).AddItem "Grouped Sub Item (t/f)"
        Me.FieldSelect(i).AddItem "Shipping - Regular"
        Me.FieldSelect(i).AddItem "Shipping - Truck Freight"
        Me.FieldSelect(i).AddItem "Shipping - Truck Freight 6.50"
        Me.FieldSelect(i).AddItem "Shipping - Free"
        Me.FieldSelect(i).AddItem "Website - Tools Plus"
        Me.FieldSelect(i).AddItem "Website - Parts"
        Me.FieldSelect(i).AddItem "EBay Price"
        Me.FieldSelect(i).AddItem "EBay Commission (from EBay Price)"
        Me.FieldSelect(i).AddItem "EBay Commission (from Web Price)"
        Me.FieldSelect(i).AddItem "EBay To Be Published (t/f)"
        Me.FieldSelect(i).AddItem "EBay Published (t/f)"
        Me.FieldSelect(i).AddItem "EBay Up (t/f)"
        Me.FieldSelect(i).AddItem "EBay Down (t/f)"
        Me.FieldSelect(i).AddItem "EBay Availability Limit"
        Me.FieldSelect(i).AddItem "EBay Best Offer (t/f)"
        Me.FieldSelect(i).AddItem "EBay Best Offer Auto-Accept"
        Me.FieldSelect(i).AddItem "EBay Category"
        Me.FieldSelect(i).AddItem "EBay Category (Store)"
        Me.FieldSelect(i).AddItem "Availability Limit (EBay)"
        
        Me.FieldSelect(i).AddItem "EBay Quantity"
        Me.FieldSelect(i).AddItem "Yahoo Quantity"
        
        Me.FieldSelect(i).AddItem "Cost < Average Cost"
        Me.FieldSelect(i).AddItem "Cost > Web Price"
        Me.FieldSelect(i).AddItem "Discontinued QOH<>0"
        Me.FieldSelect(i).AddItem "Discontinued QOH=0"
        Me.FieldSelect(i).AddItem "Discontinued QOH=0 $>0"
        Me.FieldSelect(i).AddItem "Instore=Internet (t/f)"
        Me.FieldSelect(i).AddItem "MAPP - At MAPP (t/f)"
        Me.FieldSelect(i).AddItem "MAPP - Below (t/f)"
        Me.FieldSelect(i).AddItem "Qty On S/O > Qty On Hand"
        Me.FieldSelect(i).AddItem "Qty On S/O > Qty On Hand + Qty on PO"
        Me.FieldSelect(i).AddItem "Instore>Internet (t/f)"
        Me.FieldSelect(i).AddItem "Instore<Internet (t/f)"
        
        Me.FieldSelect(i).AddItem "Kit - Any (t/f)"
        Me.FieldSelect(i).AddItem "Kit - Multi (t/f)"
        Me.FieldSelect(i).AddItem "Kit - Qty (t/f)"
        Me.FieldSelect(i).AddItem "Kit - Component (t/f)"
        
        ExpandDropDownToFit Me.FieldSelect(i)
    Next i
End Sub

Private Sub handleClauses()
    Dim clause As String
    'MsgBox Me.FieldSelect(i)
    Select Case Me.FieldSelect(i)
       
       
       'Tom's Case Statements
        Case Is = "Date Last Inventoried"
            clause = handleDate("LocationContents.LastInventoriedDate")
            
        Case Is = "Warehouse ID"
            clause = handleNumeric("LocationMaster.WarehouseID")
            
        Case Is = "Dynamic Pricing Available (t/f)"
            clause = handleText("DynamicPricing.ItemNumber")
            
        Case Is = "Dynamic Pricing Enabled (t/f)"
            clause = handleText("DynamicPricing.Enabled")
            
        Case Is = "Item Number Contains"
            clause = handleText("InventoryMaster.ItemNumber")
            
        Case Is = "Sphere 1 Category Name"
            clause = handleText("SphereOneCategories.Name")
            
        Case Is = "Product Tag Like"
            clause = handleText("ProductTags.ProductTag")
       '''''''''''''''
            
            
            
        Case Is = "Item Description"
            clause = handleText("InventoryMaster.ItemDescription")
        Case Is = "Product Line"
            clause = handleText("InventoryMaster.ProductLine")
        Case Is = "Vendor Name"
            clause = handleText("AP_Vendor.VendorName")
        Case Is = "Vendor Number"
            clause = handleText("InventoryMaster.PrimaryVendor")
        Case Is = "Additional Info"
            clause = handleText("InventoryMaster.EPN")
        Case Is = "Notes"
            clause = handleText("InventoryMaster.Components")
        'Case Is = "Weight"
        '    MsgBox "Can't handle weight!"
        '    'clause = handleWeight("InventoryMaster.Weight")
        Case Is = "Weight (any)"
            clause = handleNumeric("CAST(InventoryComponents.Weight as float)")
        Case Is = "Weight (total)"
            clause = handleNumeric("vInventoryComponentsDisplayable.TotalWeight")
        Case Is = "Length"
            clause = handleNumeric("CAST(InventoryComponents.Length as float)")
        Case Is = "Width"
            clause = handleNumeric("CAST(InventoryComponents.Width as float)")
        Case Is = "Height"
            clause = handleNumeric("CAST(InventoryComponents.Height as float)")
        'Case Is = "Dimensions"
        '    'clause = handleText("InventoryMaster.Dimensions")
        '    MsgBox "Can't handle dimensions!"
        'Case Is = "Replaced By"
        '    clause = handleMaybeNull("InventoryMaster.ReplacedBy")
        'Case Is = "Replacement For"
        '    clause = handleMaybeNull("InventoryMaster.ReplacementFor")
        Case Is = "Replaced By (t/f)"
            clause = handleMaybeEmptyBoolean("InventoryMaster.ReplacedBy")
        Case Is = "Replacement For (t/f)"
            clause = handleMaybeEmptyBoolean("InventoryMaster.ReplacementFor")
        Case Is = "Shipping Class"
            clause = handleText("InventoryMaster.Class")
        Case Is = "Discontinued"
            clause = handleBoolean("vItemStatusBreakdown.DC")
        Case Is = "Discontinued (web)"
            clause = handleBoolean("vItemStatusBreakdown.DConWeb")
        Case Is = "Being Discontinued"
            clause = handleBoolean("vItemStatusBreakdown.BeingDC")
        Case Is = "Dropship (any d/s)"
            clause = handleBoolean("vItemStatusBreakdown.DSable")
        Case Is = "Dropship Only"
            clause = handleBoolean("vItemStatusBreakdown.DSable") & Chr(0) & handleBoolean("vItemStatusBreakdown.Stock", True)
        Case Is = "Standard Pack"
            clause = handleNumeric("InventoryMaster.StdPack")
        Case Is = "Standard Cost"
            clause = handleNumeric("InventoryMaster.StdCost")
        Case Is = "Average Cost"
            clause = handleNumeric("InventoryMaster.AvgCost")
        Case Is = "Store Price"
            clause = handleNumeric("InventoryMaster.StdPrice")
        Case Is = "List Price"
            clause = handleNumeric("InventoryMaster.ListPrice")
        Case Is = "Last Price Change"
            clause = handleText("InventoryMaster.DateLastPriceChange")
        Case Is = "MAPP"
            clause = handleNumeric("InventoryMaster.MAPP")
        Case Is = "TNC"
            clause = handleNumeric("InventoryMaster.TNC")
        Case Is = "PO Comment"
            clause = handleText("InventoryMaster.POComment")
        Case Is = "New Order Quantity"
            'clause = handleNumeric("InventoryMaster.QtyOrdered")
            clause = handleNumeric("vCurrentPurchaseOrderInfo.TotalPOQuantity")
        Case Is = "Quantity On Hand"
            clause = handleNumeric("vActualWhseQty.QtyOnHand")
        Case Is = "Quantity On Sales Order"
            clause = handleNumeric("vActualWhseQty.QtyOnSO")
        Case Is = "Quantity On Purchase Order"
            clause = handleNumeric("vActualWhseQty.QtyOnPO")
        Case Is = "Quantity On Back Order"
            clause = handleNumeric("vActualWhseQty.QtyOnBO")
        Case Is = "Quantity On Hand (specific whse)"
            clause = handleWhse("InventoryQuantities.QuantityOnHand")
        Case Is = "Quantity On Sales Order (specific whse)"
            clause = handleWhse("InventoryQuantities.QuantityOnSalesOrder")
        Case Is = "Quantity On Purchase Order (specific whse)"
            clause = handleWhse("InventoryQuantities.QuantityOnPurchaseOrder")
        Case Is = "Quantity On Back Order (specific whse)"
            clause = handleWhse("InventoryQuantities.QuantityOnBackOrder")
        Case Is = "Quantity On Hand (Vendor)"
            clause = handleNumeric("InventoryMaster.VendorQuantityOnHand")
        Case Is = "Current Order Point"
        '    clause = handleNumeric("InventoryMaster.OldReorderPoint")
            clause = handleNumeric("InventoryLocationInfo.OrderPoint", "SUM")
        'Case Is = "New Order Point"
        '    clause = handleNumeric("InventoryMaster.NewReorderPoint")
        'Case Is = "Locked Order Point"
        '    clause = handleBoolean("LockOP")
        'Case Is = "Using New Order Point"
        '    clause = handleBoolean("UpdateOrderPoint")
            
        Case Is = "Sales This Month (Total)"
            clause = handleNumeric("(COALESCE(ISDAW0.Per0P2,0) + COALESCE(ISDAW0.Per0SO,0))")
        Case Is = "Sales Last Month (Total)"
            clause = handleNumeric("(COALESCE(ISDAW0.Per1P2,0) + COALESCE(ISDAW0.Per1SO,0))")
        Case Is = "Sales This Period (Total)"
            clause = handleNumeric("(COALESCE(ISDAW0.Per2P2,0) + COALESCE(ISDAW0.Per2SO,0))")
        Case Is = "Sales Last Period (Total)"
            clause = handleNumeric("(COALESCE(ISDAW0.Per3P2,0) + COALESCE(ISDAW0.Per3SO,0))")
        Case Is = "Sales Year To Date (Total)"
            clause = handleNumeric("(COALESCE(ISDAW0.Per4P2,0) + COALESCE(ISDAW0.Per4SO,0))")
        Case Is = "Sales Last Year (Total)"
            clause = handleNumeric("(COALESCE(ISDAW0.Per5P2,0) + COALESCE(ISDAW0.Per5SO,0))")
        Case Is = "Sales This Month (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per0P2,0)")
        Case Is = "Sales Last Month (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per1P2,0)")
        Case Is = "Sales This Period (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per2P2,0)")
        Case Is = "Sales Last Period (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per3P2,0)")
        Case Is = "Sales Year To Date (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per4P2,0)")
        Case Is = "Sales Last Year (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per5P2,0)")
        Case Is = "Sales This Month (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per0SO,0)")
        Case Is = "Sales Last Month (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per1SO,0)")
        Case Is = "Sales This Period (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per2SO,0)")
        Case Is = "Sales Last Period (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per3SO,0)")
        Case Is = "Sales Year To Date (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per4SO,0)")
        Case Is = "Sales Last Year (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW0.Per5SO,0)")
        
        Case Is = "Sales Rolling Last 30 (Total)"
            clause = handleNumeric("(COALESCE(ISDAW1.Per0P2,0) + COALESCE(ISDAW1.Per0SO,0))")
        Case Is = "Sales Rolling Last 30 to 60 Ago (Total)"
            clause = handleNumeric("(COALESCE(ISDAW1.Per1P2,0) + COALESCE(ISDAW1.Per1SO,0))")
        Case Is = "Sales Rolling Last 90 (Total)"
            clause = handleNumeric("(COALESCE(ISDAW1.Per2P2,0) + COALESCE(ISDAW1.Per2SO,0))")
        Case Is = "Sales Rolling Last 90 to 180 (Total)"
            clause = handleNumeric("(COALESCE(ISDAW1.Per3P2,0) + COALESCE(ISDAW1.Per3SO,0))")
        Case Is = "Sales Rolling Last 365 (Total)"
            clause = handleNumeric("(COALESCE(ISDAW1.Per4P2,0) + COALESCE(ISDAW1.Per4SO,0))")
        Case Is = "Sales Rolling Last 365 to 730 (Total)"
            clause = handleNumeric("(COALESCE(ISDAW1.Per5P2,0) + COALESCE(ISDAW1.Per5SO,0))")
        Case Is = "Sales Rolling Last 30 (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per0P2,0)")
        Case Is = "Sales Rolling Last 30 to 60 Ago (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per1P2,0)")
        Case Is = "Sales Rolling Last 90 (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per2P2,0)")
        Case Is = "Sales Rolling Last 90 to 180 (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per3P2,0)")
        Case Is = "Sales Rolling Last 365 (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per4P2,0)")
        Case Is = "Sales Rolling Last 365 to 730 (P2 Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per5P2,0)")
        Case Is = "Sales Rolling Last 30 (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per0SO,0)")
        Case Is = "Sales Rolling Last 30 to 60 (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per1SO,0)")
        Case Is = "Sales Rolling Last 90 (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per2SO,0)")
        Case Is = "Sales Rolling Last 90 to 180 (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per3SO,0)")
        Case Is = "Sales Rolling Last 365 (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per4SO,0)")
        Case Is = "Sales Rolling Last 365 to 730 (SO Only)"
            clause = handleNumeric("COALESCE(ISDAW1.Per5SO,0)")
        
        Case Is = "Returns This Month"
            clause = handleNumeric("COALESCE(ISDAW0.Per0Ret,0)")
        Case Is = "Returns Last Month"
            clause = handleNumeric("COALESCE(ISDAW0.Per1Ret,0)")
        Case Is = "Returns This Period"
            clause = handleNumeric("COALESCE(ISDAW0.Per2Ret,0)")
        Case Is = "Returns Last Period"
            clause = handleNumeric("COALESCE(ISDAW0.Per3Ret,0)")
        Case Is = "Returns Year To Date"
            clause = handleNumeric("COALESCE(ISDAW0.Per4Ret,0)")
        Case Is = "Returns Last Year"
            clause = handleNumeric("COALESCE(ISDAW0.Per5Ret,0)")
        Case Is = "Returns Rolling Last 30"
            clause = handleNumeric("COALESCE(ISDAW1.Per0Ret,0)")
        Case Is = "Returns Rolling Last 30 to 60"
            clause = handleNumeric("COALESCE(ISDAW1.Per1Ret,0)")
        Case Is = "Returns Rolling Last 90"
            clause = handleNumeric("COALESCE(ISDAW1.Per2Ret,0)")
        Case Is = "Returns Rolling Last 90 to 180"
            clause = handleNumeric("COALESCE(ISDAW1.Per3Ret,0)")
        Case Is = "Returns Rolling Last 365"
            clause = handleNumeric("COALESCE(ISDAW1.Per4Ret,0)")
        Case Is = "Returns Rolling Last 365 to 730"
            clause = handleNumeric("COALESCE(ISDAW1.Per5Ret,0)")
        
        
        Case Is = "Description 1"
            clause = handleText("PartNumbers.Desc1")
        Case Is = "Description 2"
            clause = handleText("PartNumbers.Desc2")
        Case Is = "Description 3"
            clause = handleText("PartNumbers.Desc3")
        Case Is = "Description 4"
            clause = handleText("PartNumbers.Desc4")
        
        Case Is = "Spec 1"
            clause = handleText("PartNumbers.Spec1HTML")
        Case Is = "Spec 2"
            clause = handleText("PartNumbers.Spec2HTML")
        Case Is = "Spec 3"
            clause = handleText("PartNumbers.Spec3HTML")
        Case Is = "Spec 4"
            clause = handleText("PartNumbers.Spec4HTML")
        Case Is = "Spec 5"
            clause = handleText("PartNumbers.Spec5HTML")
        Case Is = "Spec 6"
            clause = handleText("PartNumbers.Spec6HTML")
        Case Is = "Spec 7"
            clause = handleText("PartNumbers.Spec7HTML")
        Case Is = "Spec 8"
            clause = handleText("PartNumbers.Spec8HTML")
        
        Case Is = "Ext. Specs - Additional Info"
            clause = handleText("PartNumbers.ExtendedSpecsHTML", True)
        Case Is = "Ext. Specs - Write Up"
            clause = handleText("PartNumbers.WriteUpHTML", True)
        Case Is = "Ext. Specs - Features"
            clause = handleText("PartNumbers.FeaturesHTML", True)
        Case Is = "Ext. Specs - Specs"
            clause = handleText("PartNumbers.TechSpecsHTML", True)
        Case Is = "Ext. Specs - Media"
            clause = handleText("PartNumbers.MediaHTML", True)
        Case Is = "Ext. Specs - Important Info"
            clause = handleText("PartNumbers.NotesHTML", True)
        
        Case Is = "Web Orderable"
            clause = handleBoolean("PartNumbers.WebOrderable")
        Case Is = "To Be Published"
            clause = handleBoolean("PartNumbers.WebToBePublished")
        Case Is = "Published"
            clause = handleBoolean("PartNumbers.WebPublished")
        Case Is = "Web Delete?"
            clause = handleBoolean("PartNumbers.WebDelete")
        Case Is = "Web Deleted"
            clause = handleBoolean("PartNumbers.WebDeleted")
        Case Is = "Path 1"
            clause = handleWebPath("WebPaths.WebPathName", "1")
        Case Is = "Path 2"
            clause = handleWebPath("WebPaths.WebPathName", "2")
        Case Is = "Path 3"
            clause = handleWebPath("WebPaths.WebPathName", "3")
        Case Is = "Path 4"
            clause = handleWebPath("WebPaths.WebPathName", "4")
        Case Is = "Path 5"
            clause = handleWebPath("WebPaths.WebPathName", "5")
        Case Is = "Accessory"
            clause = handleText("PartNumberAccessories.Accessory")
        Case Is = "Accessory Count"
            clause = handleNumeric("PartNumberAccessories.Accessory", "COUNT")
        Case Is = "Availability"
            clause = handleText("AvailOverrides.AvailString")
        Case Is = "Availability Limit"
            clause = handleNumeric("PartNumbers.AvailabilityLimit")
        Case Is = "Sign"
            clause = handleText("SignNames.SignName")
        'Case Is = "Group Sign"
        '    clause = handleText("GroupSigns.GroupName")
        Case Is = "Special"
            clause = handleText("PartNumberSpecials.SpecialName")
        Case Is = "Section Special URL"
            clause = handleText("PartNumbers.SectionURL")
        Case Is = "Section Special Image"
            clause = handleText("PartNumbers.SectionImage")
        Case Is = "Print Sign"
            clause = handleBoolean("PartNumbers.PrintSign")
        Case Is = "Display On Front"
            clause = handleBoolean("PartNumbers.DisplayOnFront")
        Case Is = "Web Caption"
            clause = handleText("PartNumbers.WebCaption")
        Case Is = "Sales Rank (This Period)"
            clause = handleSalesRank("SalesRankings.ThisPeriod")
        Case Is = "Sales Rank (Last Period)"
            clause = handleSalesRank("SalesRankings.LastPeriod")
        Case Is = "Sales Rank (Year to Date)"
            clause = handleSalesRank("SalesRankings.YTD")
        Case Is = "Sales Rank (Last Year)"
            clause = handleSalesRank("SalesRankings.LYR")
        Case Is = "Triad Code"
            clause = handleTriadCode("InventoryMaster.TriadCode")
        'Case Is = "Truck Freight"
        '    clause = handleBoolean("InventoryMaster.TruckFreight")
        'Case Is = "Truck Freight $6.50"
        '    clause = handleBoolean("InventoryMaster.TruckFreightCheap")
        Case Is = "Reconditioned"
            clause = handleBoolean("InventoryMaster.IsReconditioned")
        Case Is = "Needs A Box"
            clause = handleBoolean("InventoryMaster.NeedsBox")
        Case Is = "Request Form Only"
            clause = handleBoolean("PartNumbers.ShowRequestForm")
        Case Is = "Pricing Critical (Non-Show)"
            clause = handlePriceCrit()
        Case Is = "Pricing Critical (Show)"
            clause = handlePriceCrit(True)
        Case Is = "Altered Item/To Be Offloaded"
            clause = handleAlteredItem()
        Case Is = "Barcode (y/n)"
            clause = handleBarcode()
        Case Is = "Status - Stock"
            clause = handleBoolean("vItemStatusBreakdown.Stock")
        Case Is = "Status - Dropship"
            clause = handleBoolean("vItemStatusBreakdown.DSable")
        Case Is = "Status - D/C by Us"
            clause = handleBoolean("vItemStatusBreakdown.DCbyUs")
        Case Is = "Status - D/C by Mfgr"
            clause = handleBoolean("vItemStatusBreakdown.DCbyMfr")
        Case Is = "GMROI"
            clause = handleNumeric("InventoryMaster.GMROI")
        Case Is = "Web/Phone Price"
            clause = handleNumeric("InventoryMaster.WebPhonePrice")
        Case Is = "Stock - Store"
            clause = handleStockInfo(1, "stock")
        Case Is = "Stock - Whse"
            clause = handleStockInfo(5, "stock")
        Case Is = "OP - Store"
            clause = handleStockInfo(1, "op")
        Case Is = "OP - Whse"
            clause = handleStockInfo(5, "op")
        Case Is = "Follies (t/f)"
            clause = handleFolliesTF("FolliesStatus")
        Case Is = "Below OP (t/f)"
            clause = handleBelowOPTF()
        Case Is = "Meta No-Index"
            clause = handleBoolean("PartNumbers.MetaNoIndex")
        Case Is = "Date Published"
            clause = handleDate("PartNumbers.PublishedDate") & Chr(0) & handleBoolean("PartNumbers.WebPublished", True)
        Case Is = "Date Reviewed"
            clause = handleDate("PartNumbers.ReviewedDate")
        Case Is = "Date Entered"
            clause = handleDate("InventoryMaster.PartIntroducedDate")
        Case Is = "EBay Availability Limit", "Availability Limit (EBay)"
            clause = handleNumeric("PartNumbers.EBayReserveQty")
        Case Is = "Profit Margin % (StdCost / Store)"
            clause = handleNumeric("InventoryMaster.StdCost", "", "<>0") & Chr(0) & _
                     handleProfitMargin("InventoryMaster.StdCost", "InventoryMaster.DiscountMarkupPriceRate1") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1", "", "<>88888.88")
        Case Is = "Profit Margin % (TNC / Store)"
            clause = handleNumeric("InventoryMaster.TNC", "", "<>0") & Chr(0) & _
                     handleProfitMargin("InventoryMaster.TNC", "InventoryMaster.DiscountMarkupPriceRate1") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1", "", "<>88888.88")
        Case Is = "Profit Margin % (AvgCost / Store)"
            clause = handleNumeric("COALESCE(InventoryMaster.AvgCost,0)", "", "<=InventoryMaster.StdCost") & Chr(0) & _
                     handleNumeric("COALESCE(InventoryMaster.AvgCost,0)", "", "<>0") & Chr(0) & _
                     handleProfitMargin("InventoryMaster.AvgCost", "InventoryMaster.DiscountMarkupPriceRate1") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1", "", "<>88888.88")
        Case Is = "Profit Margin % (StdCost / Web)"
            clause = handleNumeric("InventoryMaster.StdCost", "", "<>0") & Chr(0) & _
                     handleProfitMargin("InventoryMaster.StdCost", "InventoryMaster.IDiscountMarkupPriceRate1") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1", "", "<>0")
        Case Is = "Profit Margin % (TNC / Web)"
            clause = handleNumeric("InventoryMaster.TNC", "", "<>0") & Chr(0) & _
                     handleProfitMargin("InventoryMaster.TNC", "InventoryMaster.IDiscountMarkupPriceRate1") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1", "", "<>0")
        Case Is = "Profit Margin % (AvgCost / Web)"
            clause = handleNumeric("COALESCE(InventoryMaster.AvgCost,0)", "", "<=InventoryMaster.StdCost") & Chr(0) & _
                     handleNumeric("COALESCE(InventoryMaster.AvgCost,0)", "", "<>0") & Chr(0) & _
                     handleProfitMargin("InventoryMaster.AvgCost", "InventoryMaster.IDiscountMarkupPriceRate1") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1", "", "<>0")
        Case Is = "Profit Margin $ (StdCost / Store)"
            clause = handleNumeric("InventoryMaster.StdCost", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1", "", "<>88888.88") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1-InventoryMaster.StdCost")
        Case Is = "Profit Margin $ (TNC / Store)"
            clause = handleNumeric("InventoryMaster.TNC", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1", "", "<>88888.88") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1-InventoryMaster.TNC")
        Case Is = "Profit Margin $ (AvgCost / Store)"
            clause = handleNumeric("COALESCE(InventoryMaster.AvgCost,0)", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1", "", "<>88888.88") & Chr(0) & _
                     handleNumeric("InventoryMaster.DiscountMarkupPriceRate1-COALESCE(InventoryMaster.AvgCost,0)")
        Case Is = "Profit Margin $ (StdCost / Web)"
            clause = handleNumeric("InventoryMaster.StdCost", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1-InventoryMaster.StdCost")
        Case Is = "Profit Margin $ (TNC / Web)"
            clause = handleNumeric("InventoryMaster.TNC", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1-InventoryMaster.TNC")
        Case Is = "Profit Margin $ (AvgCost / Web)"
            clause = handleNumeric("COALESCE(InventoryMaster.AvgCost,0)", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1", "", "<>0") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1-COALESCE(InventoryMaster.AvgCost,0)")
        
        Case Is = "Grouped Master Item (t/f)"
            clause = handleBoolean("PartNumbers.WebPointered")
        Case Is = "Grouped Sub Item (t/f)"
            clause = handleGroupSub()
        Case Is = "Shipping - Regular"
            clause = handleIndexedBoolean("InventoryMaster.ShippingType", 0)
        Case Is = "Shipping - Truck Freight"
            clause = handleIndexedBoolean("InventoryMaster.ShippingType", 1)
        Case Is = "Shipping - Truck Freight 6.50"
            clause = handleIndexedBoolean("InventoryMaster.ShippingType", 2)
        Case Is = "Shipping - Free"
            clause = handleIndexedBoolean("InventoryMaster.ShippingType", 3)
        Case Is = "Website - Tools Plus"
            clause = handleIndexedBoolean("InventoryMaster.WebsiteID", 0)
        Case Is = "Website - Parts"
            clause = handleIndexedBoolean("InventoryMaster.WebsiteID", 1)
        Case Is = "EBay Price"
            clause = handleNumeric("InventoryMaster.EBayPrice")
        Case Is = "EBay Commission (from EBay Price)"
            clause = handleNumeric("dbo.CalcEBayCommission(InventoryMaster.EBayPrice)")
        Case Is = "EBay Commission (from Web Price)"
            clause = handleNumeric("dbo.CalcEBayCommission(InventoryMaster.IDiscountMarkupPriceRate1)")
        Case Is = "EBay To Be Published (t/f)"
            clause = handleBoolean("PartNumbers.EBayToBePublished")
        Case Is = "EBay Published (t/f)"
            clause = handleBoolean("PartNumbers.EBayPublished")
        Case Is = "EBay Up (t/f)"
            clause = handleIndexedBoolean("PartNumbers.EBayStatusID", 1)
        Case Is = "EBay Down (t/f)"
            clause = handleIndexedBoolean("PartNumbers.EBayStatusID", 0)
        Case Is = "EBay Best Offer (t/f)"
            clause = handleBoolean("InventoryMaster.EBayAllowBestOffer")
        Case Is = "EBay Best Offer Auto-Accept"
            clause = handleNumeric("InventoryMaster.EBayBestOfferAutoAccept")
        Case Is = "EBay Category"
            clause = handleText("EBayCategoryE.CategoryName")
        Case Is = "EBay Category (Store)"
            clause = handleText("EBayCategoryS.CategoryName")
            
        Case Is = "EBay Quantity"
            clause = handleNumeric("vActualWhseQty.QtyOnHand-vActualWhseQty.QtyOnSO-PartNumbers.EBayReserveQty-PartNumbers.AvailabilityLimit")
        Case Is = "Yahoo Quantity"
            clause = handleNumeric("vActualWhseQty.QtyOnHand-vActualWhseQty.QtyOnSO-PartNumbers.AvailabilityLimit")
        
        Case Is = "Cost < Average Cost"
            clause = handleDirectClause("InventoryMaster.StdCost+0.5<InventoryMaster.AvgCost") & Chr(0) & _
                     handleDirectClause("vItemStatusBreakdown.DC=0")
        Case Is = "Cost > Web Price"
            clause = handleDirectClause("(InventoryMaster.StdCost>=InventoryMaster.IDiscountMarkupPriceRate1 OR InventoryMaster.TNC>=InventoryMaster.IDiscountMarkupPriceRate1)") & Chr(0) & _
                     handleDirectClause("InventoryMaster.IDiscountMarkupPriceRate1<>0") & Chr(0) & _
                     handleDirectClause("vItemStatusBreakdown.DC=0") & Chr(0) & _
                     handleDirectClause("PartNumbers.WebPublished=1")
        Case Is = "Discontinued QOH<>0"
            clause = handleDirectClause("vItemStatusBreakdown.DC=1") & Chr(0) & _
                     handleDirectClause("vActualWhseQty.QtyOnHand<>0")
        Case Is = "Discontinued QOH=0"
            clause = handleDirectClause("vItemStatusBreakdown.DC=1") & Chr(0) & _
                     handleDirectClause("vActualWhseQty.QtyOnHand=0")
        Case Is = "Discontinued QOH=0 $>0"
            clause = handleDirectClause("vItemStatusBreakdown.DC=1") & Chr(0) & _
                     handleDirectClause("vActualWhseQty.QtyOnHand=0") & Chr(0) & _
                     handleDirectClause("InventoryMaster.IDiscountMarkupPriceRate1<>0")
        Case Is = "Instore=Internet (t/f)"
            clause = handleFieldComparison("InventoryMaster.DiscountMarkupPriceRate1", "InventoryMaster.IDiscountMarkupPriceRate1", "=") & Chr(0) & _
                     handleDirectClause("vItemStatusBreakdown.DC=0")
        Case Is = "MAPP - At MAPP (t/f)"
            clause = handleDirectClause("InventoryMaster.MAPP<>0") & Chr(0) & _
                     handleFieldComparison("InventoryMaster.MAPP", "InventoryMaster.IDiscountMarkupPriceRate1", "=") & Chr(0) & _
                     handleDirectClause("vItemStatusBreakdown.DC=0")
        Case Is = "MAPP - Below (t/f)"
            clause = handleDirectClause("InventoryMaster.MAPP<>0") & Chr(0) & _
                     handleFieldComparison("InventoryMaster.IDiscountMarkupPriceRate1", "InventoryMaster.MAPP", "<") & Chr(0) & _
                     handleDirectClause("vItemStatusBreakdown.DC=0")
        Case Is = "Qty On S/O > Qty On Hand"
            clause = handleFieldComparison("vActualWhseQty.QtyOnSO", "vActualWhseQty.QtyOnHand", ">") & Chr(0) & _
                     handleDirectClause("vActualWhseQty.QtyOnSO>0")
        Case Is = "Qty On S/O > Qty On Hand + Qty on PO"
            clause = handleFieldComparison("vActualWhseQty.QtyOnSO", "vActualWhseQty.QtyOnHand+vActualWhseQty.QtyOnPO", ">") & Chr(0) & _
                     handleDirectClause("vActualWhseQty.QtyOnSO>0")
        Case Is = "Instore>Internet (t/f)"
            clause = handleNumeric("InventoryMaster.DiscountMarkupPriceRate1", "", "<>88888.88") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1", "", "<>0") & Chr(0) & _
                     handleFieldComparison("InventoryMaster.DiscountMarkupPriceRate1", "InventoryMaster.IDiscountMarkupPriceRate1", ">")
        Case Is = "Instore<Internet (t/f)"
            clause = handleNumeric("InventoryMaster.DiscountMarkupPriceRate1", "", "<>88888.88") & Chr(0) & _
                     handleNumeric("InventoryMaster.IDiscountMarkupPriceRate1", "", "<>0") & Chr(0) & _
                     handleFieldComparison("InventoryMaster.DiscountMarkupPriceRate1", "InventoryMaster.IDiscountMarkupPriceRate1", "<")
        Case Is = "Kit - Any (t/f)"
            clause = handleBoolean("InventoryMaster.IsMASKit")
        Case Is = "Kit - Multi (t/f)"
            clause = handleBoolean("InventoryMaster.IsMASKit") & Chr(0) & _
                     handleDirectClause("1<(SELECT COUNT(*) FROM IM_SalesKitDetail WHERE SalesKitNo=InventoryMaster.ItemNumber)")
        Case Is = "Kit - Qty (t/f)"
            clause = handleBoolean("InventoryMaster.IsMASKit") & Chr(0) & _
                     handleDirectClause("1=(SELECT COUNT(*) FROM IM_SalesKitDetail WHERE SalesKitNo=InventoryMaster.ItemNumber)")
        Case Is = "Kit - Component (t/f)"
            clause = handleDirectClause("0<(SELECT COUNT(*) FROM IM_SalesKitDetail WHERE ComponentItemCode=InventoryMaster.ItemNumber)")
            
        'Tom's Code 9-24-2014 to add filter for subtitled items... dont need i dont think...
        Case Is = "EBay Out Of Stock (t/f)"
            If Me.TextToMatch(i) = "t" Or Me.TextToMatch(i) = "T" Then
                'clause = handleDirectClause("PartNumbers.ItemNumber=InventoryMaster.ItemNumber AND EBayPublished=1 AND EBayStatusID=1 AND EBayAvailableQty < 1")
                clause = handleDirectClause("InventoryMaster.ItemNumber=vActualWhseQty.ItemNumber AND QtyOnHand-QtyOnSO-QtyOnBO < 1 AND PartNumbers.EBayPublished=1 AND PartNumbers.EBayStatusID=1 AND PartNumbers.OutOfStock=1")
            End If
            If Me.TextToMatch(i) = "f" Or Me.TextToMatch(i) = "F" Then
                'clause = handleDirectClause("PartNumbers.ItemNumber=InventoryMaster.ItemNumber AND EBayPublished=1 AND EBayStatusID=1 AND EBayAvailableQty > 0")
                clause = handleDirectClause("InventoryMaster.ItemNumber=vActualWhseQty.ItemNumber AND QtyOnHand-QtyOnSO-QtyOnBO >= 1 AND PartNumbers.EBayPublished=1 AND PartNumbers.EBayStatusID=1 AND OutOfStock=0")
            End If
        Case Is = "Subtitle Overridden Item (t/f)"
            If Me.TextToMatch(i) = "t" Or Me.TextToMatch(i) = "T" Then
            clause = handleDirectClause("Subtitle Overridden Item (t/f)")
            End If
            If Me.TextToMatch(i) = "f" Or Me.TextToMatch(i) = "F" Then
            clause = handleDirectClause("Subtitle Overridden Item (t/f) InventoryMaster.ItemNumber NOT IN (SELECT ItemNumber FROM SubtitledItems)")
            End If
        Case Is = "Subtitled By Default (t/f)"
            If Me.TextToMatch(i) = "t" Or Me.TextToMatch(i) = "T" Then
            clause = handleDirectClause("Subtitled By Default (true)")
            End If
            If Me.TextToMatch(i) = "f" Or Me.TextToMatch(i) = "F" Then
            clause = handleDirectClause("Subtitled By Default (false)")
            End If
        Case Is = "Subtitled (t/f)"
            If Me.TextToMatch(i) = "t" Or Me.TextToMatch(i) = "T" Then
            clause = handleDirectClause("Subtitled (true)")
            End If
            If Me.TextToMatch(i) = "f" Or Me.TextToMatch(i) = "F" Then
            clause = handleDirectClause("Subtitled (false)")
            End If
            
        Case Else
            If InStr(Me.FieldSelect(i), "In-Store Price Break") Then
                clause = handleNumeric("InventoryMaster.DiscountMarkupPriceRate" & Right$(Me.FieldSelect(i), 1))
            ElseIf InStr(Me.FieldSelect(i), "In-Store Quantity Break") Then
                clause = handleNumeric("InventoryMaster.BreakQty" & Right$(Me.FieldSelect(i), 1))
            ElseIf InStr(Me.FieldSelect(i), "Internet Price Break") Then
                clause = handleNumeric("InventoryMaster.IDiscountMarkupPriceRate" & Right$(Me.FieldSelect(i), 1))
            ElseIf InStr(Me.FieldSelect(i), "Internet Quantity Break") Then
                clause = handleNumeric("InventoryMaster.IBreakQty" & Right$(Me.FieldSelect(i), 1))
            End If
    
    End Select
    
    If InStr(clause, Chr(0)) Then
        clause = "(" & Join(Split(clause, Chr(0)), " AND ") & ")"
    End If
    
    If i <> 0 Then
        If cType = "where" Then
            whereClause = whereClause & IIf(whereClause = "", "", " " & Me.AndOr(i - 1) & " ") & clause
        Else
            havingClause = havingClause & IIf(havingClause = "", "", " " & Me.AndOr(i - 1) & " ") & clause
        End If
    Else
        If cType = "where" Then
            whereClause = clause
        Else
            havingClause = clause
        End If
    End If
End Sub

Private Function handleText(tableField As String, Optional longText As Boolean = False) As String
    cType = "where"
    If Me.TextToMatch(i) = "IS NULL" Or Me.TextToMatch(i) = "IS NOT NULL" Then
        handleText = tableField & " " & Me.TextToMatch(i)
    ElseIf Me.TextToMatch(i) = "IS EMPTY" Then
        If longText Then
            handleText = "(DATALENGTH(" & tableField & ") = 0 OR " & tableField & " IS NULL)"
        Else
            handleText = "(" & tableField & " = '' OR " & tableField & " IS NULL)"
        End If
    ElseIf Me.TextToMatch(i) = "IS NOT EMPTY" Then
        If longText Then
            handleText = "DATALENGTH(" & tableField & ") >0 AND " & tableField & " IS NOT NULL"
        Else
            handleText = tableField & " <> '' AND " & tableField & " IS NOT NULL"
        End If
    ElseIf Left(Me.TextToMatch(i), 8) = "EXACTLY " Then
        handleText = tableField & " = '" & EscapeSQuotes(Mid(Me.TextToMatch(i), 9)) & "'"
    ElseIf Left(Me.TextToMatch(i), 4) = "NOT " Then
        handleText = "( " & tableField & " NOT LIKE '%" & EscapeSQuotes(Mid(Me.TextToMatch(i), 5)) & "%' OR " & tableField & " IS NULL )"
    ElseIf Left(Me.TextToMatch(i), 11) = "BEGINSWITH " Then
        handleText = tableField & " LIKE '" & EscapeSQuotes(Mid(Me.TextToMatch(i), 12)) & "%'"
    ElseIf Left(Me.TextToMatch(i), 9) = "ENDSWITH " Then
        handleText = tableField & " LIKE '%" & EscapeSQuotes(Mid(Me.TextToMatch(i), 10)) & "'"
    
    Else
        handleText = tableField & " LIKE '%" & EscapeSQuotes(Me.TextToMatch(i)) & "%'"
    End If
End Function

Private Function handleBoolean(tableField As String, Optional invert As Boolean = False) As String
    cType = "where"
    If Not invert Then
        handleBoolean = tableField & "=" & IIf(scanForBoolean(), 1, 0)
    Else
        handleBoolean = tableField & "=" & IIf(scanForBoolean(), 0, 1)
    End If
End Function

'Private Function handleMaybeNull(tableField As String) As String
'    cType = "where"
'    handleMaybeNull = "(" & tableField & " IS " & IIf(scanForBoolean(), "NOT ", "") & "NULL" & " OR " & _
'                            tableField & IIf(scanForBoolean(), "<>", "=") & "''" & ")"
'End Function

Private Function handleMaybeEmptyBoolean(tableField As String) As String
    cType = "where"
    Dim negate As Boolean
    negate = scanForBoolean() 'true means not empty, so false?
    handleMaybeEmptyBoolean = "(" & tableField & " IS " & IIf(negate, "NOT ", "") & "NULL" & IIf(negate, " AND ", " OR ") & _
                                    tableField & IIf(negate, "<>", "=") & "''" & ")"
End Function

Private Function handleNumeric(tableField As String, Optional aggregate As String = "", Optional override As String = "") As String
    If aggregate = "" Then
        cType = "where"
        If LCase(Me.TextToMatch(i)) = "availabilitylimit" Then
            'handleNumeric = tableField & IIf(override = "", scanForComparison() & "'PartNumbers.AvailabilityLimit'", override)
            If override = "" Then
                handleNumeric = tableField & scanForComparison() & "'PartNumbers.AvailabilityLimit'"
            Else
                handleNumeric = tableField & override
            End If
        Else
            'handleNumeric = tableField & IIf(override = "", scanForComparison() & Me.TextToMatch(i), override)
            If override = "" Then
                handleNumeric = tableField & scanForComparison() & Me.TextToMatch(i)
            Else
                handleNumeric = tableField & override
            End If
        End If
    Else
        cType = "having"
        handleNumeric = aggregate & "(" & tableField & ")" & IIf(override = "", scanForComparison & Me.TextToMatch(i), override)
    End If
End Function

Private Function handleDate(tableField As String) As String
    cType = "where"
    Dim comparator As String
    comparator = scanForComparison()
    If comparator = "=" Then
        MsgBox "Equals comparison for a date probably doesn't do what you think! (you should really use less than/greater than) Continuing anyway!"
    End If
    handleDate = tableField & comparator & "'" & Me.TextToMatch(i) & "'"
End Function

Private Function handleWebPath(tableField As String, pathLevel As String) As String
    cType = "where"
    handleWebPath = "EXISTS (SELECT ItemNumber FROM PartNumberWebPaths INNER JOIN WebPaths ON PartNumberWebPaths.WebPathID=WebPaths.ID WHERE PartNumbers.ItemNumber=PartNumberWebPaths.ItemNumber AND " & handleText(tableField) & " AND WebPaths.PathLevel=" & pathLevel & ")"
End Function

Private Function handleWhse(tableField As String) As String
    cType = "where"
    Dim whse As String
    Dim amount As String
    If Not CBool(InStr(Me.TextToMatch(i), " ")) Then
        MsgBox "Error, invalid parameter for " & Me.FieldSelect(i) & ", skipping this filter"
        handleWhse = "1=1"
    End If
    whse = Left(Me.TextToMatch(i), InStr(Me.TextToMatch(i), " ") - 1)
    amount = Mid(Me.TextToMatch(i), InStr(Me.TextToMatch(i), " ") + 1)
    Me.TextToMatch(i) = amount
    handleWhse = "InventoryQuantities.WhseCode='" & whse & "' AND " & handleNumeric(tableField)
End Function

Private Function handleSalesRank(period As String) As String
    cType = "where"
    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = "^(dollars|units) ([a-d])$"
    re.IgnoreCase = True
    If re.test(Me.TextToMatch(i)) Then
        handleSalesRank = period & Left(Me.TextToMatch(i), InStr(Me.TextToMatch(i), " ") - 1) & "SalesRank='" & Right(Me.TextToMatch(i), 1) & "'"
    Else
        MsgBox "Can't figure out sales rank! Should be something like ""dollars A"" or ""units C"""
        handleSalesRank = "1=1"
    End If
End Function

'Private Function handleWeight(tableField As String) As String
'    cType = "where"
'    If InStr(Me.TextToMatch(i), ";") Then
'        handleWeight = handleText(tableField) 'probably just a search for multi-box items?
'    Else
'        handleWeight = handleNumeric("TmpWeight.WeightNumeric")
'    End If
'End Function
'Private Function handleWeight(summation As Boolean) As String
'    cType = "where"
'    If summation Then
'        handleWeight = handleNumeric("vInventoryComponentsDisplayable.TotalWeight")
'    Else
'        handleWeight = handleNumeric("CAST(InventoryComponents.Weight as float)")
'    End If
'End Function

Private Function handleTriadCode(tableField As String) As String
    cType = "where"
    If Len(Me.TextToMatch(i)) = 1 Or Len(Me.TextToMatch(i)) = 2 Then
        Me.TextToMatch(i) = Me.TextToMatch(i) & "  "
    End If
    Dim pos As Long, retval As String
    For pos = 1 To 3
        If Mid(Me.TextToMatch(i), pos, 1) <> " " Then
            retval = IIf(retval = "", "", retval & " AND ") & "SUBSTRING(TriadCode," & pos & ",1)='" & Mid(Me.TextToMatch(i), pos, 1) & "'"
        End If
    Next pos
    handleTriadCode = retval
End Function

Private Function handlePriceCrit(Optional saleDisc As Boolean = False) As String
    cType = "where"
    Dim retval As String
    retval = "("
    retval = retval & "(InventoryMaster.StdCost>InventoryMaster.TNC AND InventoryMaster.StdCost>InventoryMaster.IDiscountMarkupPriceRate1" & IIf(saleDisc = True, "*0.95", "") & ")"
    retval = retval & " OR "
    retval = retval & "(InventoryMaster.StdCost<=InventoryMaster.TNC AND InventoryMaster.TNC>InventoryMaster.IDiscountMarkupPriceRate1" & IIf(saleDisc = True, "*0.95", "") & ")"
    retval = retval & ")"
    retval = retval & " AND vItemStatusBreakdown.DC=0"
    retval = retval & " AND vItemStatusBreakdown.BeingDC=0"
    handlePriceCrit = retval
End Function

Private Function handleAlteredItem() As String
    cType = "where"
    If scanForBoolean() Then
        handleAlteredItem = "PartNumbers.LastChanged>=PartNumbers.LastOffload AND PartNumbers.WebPublished=1"
    Else
        handleAlteredItem = "PartNumbers.LastChanged<PartNumbers.LastOffload AND PartNumbers.WebPublished=1"
    End If
End Function

Private Function handleBarcode() As String
    cType = "having"
    handleBarcode = "COUNT(vBarcodes.Barcode)" & IIf(scanForBoolean(), ">=1", "=0")
End Function

Private Function handleStockInfo(warehouseID As Long, whichField As String) As String
    cType = "where"
    Dim retval As String
    retval = "(InventoryLocationInfo.WarehouseID=" & warehouseID & " AND "
    If whichField = "stock" Then
        retval = retval & handleBoolean("InventoryLocationInfo.StockHere")
    ElseIf whichField = "op" Then
        retval = retval & handleNumeric("InventoryLocationInfo.OrderPoint")
    End If
    retval = retval & ")"
    
    handleStockInfo = retval
End Function

Private Function handleFolliesTF(tableField As String, Optional invert As Boolean = False) As String
    cType = "where"
    If scanForBoolean() Then
        handleFolliesTF = tableField & IIf(invert, "=", "<>") & "0"
    Else
        handleFolliesTF = tableField & IIf(invert, "<>", "=") & "0"
    End If
End Function

Private Function handleBelowOPTF() As String
    cType = "where"
    handleBelowOPTF = "vActualWhseQty.QtyOnHand" & IIf(scanForBoolean(), "<", ">=") & "vInventoryLocationInfoSummed.TotalOrderPoint-vActualWhseQty.QtyOnPO"
End Function

Private Function handleProfitMargin(tableFieldCost As String, tableFieldPrice) As String
    cType = "where"
    handleProfitMargin = tableFieldPrice & "/" & tableFieldCost & "*100-100" & scanForComparison() & Me.TextToMatch(i)
End Function

Private Function handleGroupSub() As String
    cType = "where"
    handleGroupSub = "PartNumberGroupLines.ItemNumber IS " & IIf(scanForBoolean(), "NOT ", "") & "NULL"
End Function

Private Function handleIndexedBoolean(tableField As String, val As Long, Optional invert As Boolean = False) As String
    cType = "where"
    If scanForBoolean() Then
        handleIndexedBoolean = tableField & IIf(invert, "<>", "=") & val
    Else
        handleIndexedBoolean = tableField & IIf(invert, "=", "<>") & val
    End If
End Function

Private Function handleDirectClause(clause As String) As String
    cType = "where"
    handleDirectClause = clause
End Function

Private Function handleFieldComparison(tableField1 As String, tableField2 As String, comparator As String) As String
    cType = "where"
    Dim negate As Boolean
    negate = Not scanForBoolean()
    handleFieldComparison = IIf(negate, "NOT(", "") & tableField1 & comparator & tableField2 & IIf(negate, ")", "")
End Function

Private Function scanForComparison() As String
    If Left(Me.TextToMatch(i), 2) = "<=" Or Left(Me.TextToMatch(i), 2) = ">=" Or Left(Me.TextToMatch(i), 2) = "<>" Then
        scanForComparison = Left(Me.TextToMatch(i), 2)
        Me.TextToMatch(i) = Mid(Me.TextToMatch(i), 3)
    ElseIf Left(Me.TextToMatch(i), 1) = "<" Or Left(Me.TextToMatch(i), 1) = ">" Or Left(Me.TextToMatch(i), 1) = "=" Then
        scanForComparison = Left(Me.TextToMatch(i), 1)
        Me.TextToMatch(i) = Mid(Me.TextToMatch(i), 2)
    Else
        scanForComparison = "="
    End If
End Function

Private Function scanForBoolean() As Boolean
    scanForBoolean = CBool(UCase(Left(Me.TextToMatch(i), 1)) = "Y" Or UCase(Left(Me.TextToMatch(i), 1)) = "T")
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not fromOK Then
        If Me.hiddenOwner = "InventoryMaintenance" Then
            InventoryMaintenance.btnFilterByForm.tag = ""
        End If
        If Me.hiddenOwner = "SignMaintenance" Then
            SignMaintenance.btnFilterByForm.tag = ""
        End If
    End If
End Sub



