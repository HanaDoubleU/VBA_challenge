Sub macro()

    ' ---
    ' labels
    ' ---

    ' declaring variables
    ' source: solutions' images from monday's lecture
    Dim ipLabel As String
    Dim jLabel As String
    Dim kLabel As String
    Dim lLabel As String
    Dim o1stLabel As String
    Dim o2ndLabel As String
    Dim o3rdLabel As String
    Dim qLabel As String
    
    ' assigning values to variables
    'source: solutions' images from monday's lecture
    ipLabel = "<ticker>"
    jLabel = "<quarterly change>"
    kLabel = "<percent change>"
    lLabel = "<total stock volume>"
    o1stLabel = "<greatest % increase>"
    o2ndLabel = "<greatest % decrease>"
    o3rdLabel = "<greatest total volume>"
    qLabel = "<value>"
    
    ' adding labels
    ' source: solutions' images from monday's lecture
    Cells(1, 9).Value = ipLabel
    Cells(1, 10).Value = jLabel
    Cells(1, 11).Value = kLabel
    Cells(1, 12).Value = lLabel
    Cells(2, 15).Value = o1stLabel
    Cells(3, 15).Value = o2ndLabel
    Cells(4, 15).Value = o3rdLabel
    Cells(1, 16).Value = ipLabel
    Cells(1, 17).Value = qLabel
    
    ' assigning value to sheet
    ' source: solutions' images from thursday's lecture
    Set first_ws = Worksheets("A")

    ' autofitting columns
    ' source: solutions' images from thursday's lecture
    first_ws.Columns("I:Q").AutoFit

    ' ---
    ' loop
    ' ---

    ' declaring variables
    ' source: solutions' images from monday's lecture
    Dim t As String
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim iLastRow As Long

    ' assigning value to variable
    ' source: solutions' images from monday's lecture
    t = "AAB"
    LastRow = 0
    iLastRow = 0

    ' (1) assigning value to variable, and (2) looping through sheets
    ' source: (1) tutoring sessions and xpert, and (2) solutions' images from thursday's lecture
    For Each ws in Worksheets

        ' re-assigning value to variable before looping through each sheet
        ' source: (1) tutoring sessions and xpert, and (2) solutions' images from thursday's lecture
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' looping through rows
        ' source: solutions' images from tuesday and thursday's lectures
        For i = 2 to LastRow

            ' re-assigning value to variable before each paste
            ' source: (1) tutoring sessions and xpert, and (2) solutions' images from thursday's lecture
            iLastRow = first_ws.Cells(Rows.Count, "I").End(xlUp).Row + 1

                ' copying first symbol under <ticker> to new column
                ' (1) solutions' images from monday's lecture, and (2) xpert
                If ws.Cells(i+1, 1).Value <> t Then
                first_ws.Cells(iLastRow, 9).Value = t
               
                ' no more conditional
                End If

            ' re-assigning value to variable after conditional while looping
            ' source: tutoring sessions and xpert
            t = ws.Cells(i + 1, 1).Value

        ' no more looping through rows
        Next i

    ' no more looping through sheets
    Next ws

End Sub