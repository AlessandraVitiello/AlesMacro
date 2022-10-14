Attribute VB_Name = "Module2"
Sub GEAOpenOrders()
Attribute GEAOpenOrders.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GEAOpenOrders Macro
'
' INSTRUCTIONS:
'
' 1. Create a folder for this process. You will always use this same folder.
'
' 2. Folder structure:
' There should be two files in this folder:
'   * Your report
'   * A text file named FilterNumbers.txt containing the BPID's you want to check against.
'
' 3. Change the path to your folder on the line further down
'
' You are ready to start the macro!






    Range("A:A,D:D,H:H,I:I,J:J").Select
    Range("J1").Activate
    Range("A:A,D:D,H:H,I:I,J:J,Q:Q").Select
    Range("Q1").Activate
    Range("A:A,D:D,H:H,I:I,J:J,Q:Q,U:U,V:V").Select
    Range("V1").Activate
    Range( _
        "A:A,D:D,H:H,I:I,J:J,Q:Q,U:U,V:V,AE:AE,AF:AF,AG:AG,AH:AH,AI:AI,AJ:AJ,AK:AK,AL:AL,AM:AM,AN:AN,AO:AO,AP:AP" _
        ).Select
    Range("AP1").Activate

    Range( _
        "A:A,D:D,H:H,I:I,J:J,Q:Q,U:U,V:V,AE:AE,AF:AF,AG:AG,AH:AH,AI:AI,AJ:AJ,AK:AK,AL:AL,AM:AM,AN:AN,AO:AO,AP:AP,AQ:AQ,AR:AR,AS:AS,AT:AT,AU:AU,AW:AW" _
        ).Select
    Range("AW1").Activate
    Selection.Delete Shift:=xlToLeft


    
     
    Dim my_file As Integer
    Dim text_line As String
    Dim file_name As String
    Dim i As Integer
    
' CHANGE FILEPATH HERE
' =======================================================================
' -----------------------------------------------------------------------

    file_name = "C:\Users\jonas\Downloads\Excel\FilterNumbers.txt"
    
'------------------------------------------------------------------------
'========================================================================

    my_file = FreeFile()
    Open file_name For Input As my_file

    i = 1
    
    Dim myArray() As Variant
    ReDim myArray(1)
   

    While Not EOF(my_file)
        Line Input #my_file, text_line
        Cells(i, "AA").Value = text_line
        ReDim Preserve myArray(i)
        myArray(i) = text_line
        i = i + 1
   
    Wend
    
    ActiveSheet.Range("$A$1:$Y$335665").AutoFilter Field:=14, Criteria1:=myArray, Operator:=xlFilterValues
    
End Sub


