    Option Explicit
    
    ' general inputs
    Dim genInput As String
    Dim cancelled As Boolean
    Dim matInput As String
    Dim hdiamInput As String
    Dim hlengthInput As String
    Dim hwidthInput As String
    Dim punchInput As String
    Dim quantityInput As String
    Dim hsInput As String
    Dim sordernum As String
    Dim cordername As String
    Dim savename As String
    Dim longstatus As Long, longwarnings As Long
    Dim rakenyc As Integer
    Dim rakeInput As String
    
    ' user stated inputs
    Dim sheight As Single
    Dim swidth As Single
    Dim length As Single
    Dim thickness As Single
    Dim hprofile As String
    Dim hlength As Single
    Dim hwidth As Single
    Dim hdiam As Single
    Dim punch As Integer
    Dim mat As Integer
    Dim cdist As Single
    Dim quantity As Integer
    Dim holeside As Integer
    Dim rlength As Single
    Dim flength As Single
    Dim fclength As Single
    Dim rake As Single
    
    ' hole calcs variables
    Dim hnumraw As Integer
    Dim hnum As Integer
    Dim hnumFinput As String
    Dim hnumF As Integer
    Dim gsize As Single
    Dim egapraw As Single
    Dim egap As Single
    Dim egap40 As Single
    Dim egap2 As Single
    Dim egapF As Single
    Dim cgap As Integer
    Dim egleft As Single
    Dim egleftinput As String
    Dim egright As Single
    Dim clength As Single
    Dim clengthnyc As Integer
    Dim clengthinput As String
    
    Sub laserm8_variable_intros()
        Dim swApp As SldWorks.SldWorks
        Dim swModel As SldWorks.ModelDoc2
        Set swApp = Application.SldWorks
        Set swModel = swApp.ActiveDoc
    
        ' Display a message box with OK and Cancel buttons
        Dim result2 As VbMsgBoxResult
        result2 = MsgBox("Welcome to the LASERM8 V11.2 Beta Interface !" & vbCrLf & "Last update: 22/02/2024" & vbCrLf & vbCrLf & "! WARNING !" _
        & vbCrLf & "Beta version contains some bugs that are currently being fixed." & vbCrLf & vbCrLf & "Please press OK to continue", vbOKCancel + vbInformation, "FC Racing 2024 " & Chr(169))
        
        ' Check the result
        If result2 = vbCancel Then
            ' User clicked Cancel, perform cancellation actions
            MsgBox "Input cancelled. Exiting.", vbOKOnly, "FC Racing 2024 " & Chr(169)
            Exit Sub ' Exit the subroutine or perform other cancellation actions
        End If
        
        ' known bugs update
        Dim result3 As VbMsgBoxResult
        result3 = MsgBox("Known bugs:" & vbCrLf & vbCrLf & "1) Raked products will sometimes not generate properly " _
        & "when on large rake angles, small section widths and small holes." _
        & vbCrLf & vbCrLf & "If this happens, take note of the final data before construction and input manually.", vbOKCancel + vbExclamation, "FC Racing 2024" & Chr(169))
        
        ' Check the result
        If result3 = vbCancel Then
            ' User clicked Cancel, perform cancellation actions
            MsgBox "Input cancelled. Exiting.", vbOKOnly, "FC Racing 2024 " & Chr(169)
            Exit Sub ' Exit the subroutine or perform other cancellation actions
        End If
        
        ' Validate and get set parameters (sheight, swidth, thickness and length)
        If Not SetParameters("section dimension height (the longest side)", sheight) Then Exit Sub
        If Not SetParameters("section dimension width (the shortest side)", swidth) Then Exit Sub
        If Not SetParameters("thickness", thickness) Then Exit Sub
        If Not SetParameters("overall length", length) Then Exit Sub
        
        ' Validate and get material
        Do
        matInput = UCase(InputBox("Please enter A for ALU or G for GAL or S for STAINLESS STEEL (SS) or P for PAINTED.", "FC Racing 2024 " & Chr(169)))
    
            If matInput = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
            Exit Sub ' Exit the function if the user canceled
            
            ElseIf Not (matInput Like "A" Or matInput Like "G" Or matInput Like "S" Or matInput Like "P") Then
                cancelled = False
                MsgBox "Please enter either the letter A, G, S or P only."
            End If
        Loop Until matInput Like "A" Or matInput Like "G" Or matInput Like "S" Or matInput Like "P" Or cancelled
    
        ' Convert the input to an integer
        'mat = CInt(matInput)
        
        ' Validate and get hole profile
        Do
            hprofile = UCase(InputBox("Please select a hole profile from the following:" & vbCrLf & _
                "C = Circle, S = Square or R = Rectangle", "FC Racing 2024 " & Chr(169)))
            If hprofile = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
            Exit Sub ' Exit the function if the user cancelled
            
            ElseIf Not (hprofile Like "C" Or hprofile Like "S" Or hprofile Like "R") Then
                cancelled = False
                MsgBox "Please input C, S, or R only."
            End If
        Loop Until hprofile Like "C" Or hprofile Like "S" Or hprofile Like "R" Or cancelled
    
        ' Handle different hole profiles
        Select Case hprofile
            Case "C"
                ' Circle profile, prompt for diameter
                Do
                hdiamInput = InputBox("Please enter the diameter for the circle profile.", "FC Racing 2024 " & Chr(169))
                If hdiamInput = "" Then
                    cancelled = True
                    MsgBox "Input cancelled. Exiting."
                    Exit Sub ' Exit the function if the user cancelled
                ElseIf IsNumeric(hdiamInput) Then
                    hdiam = CSng(hdiamInput)
                        If hdiam < 0 Then
                            cancelled = False
                            MsgBox "Please enter a positive value for the diameter."
                        ElseIf hdiam = 0 Then
                            cancelled = False
                            MsgBox "Please enter a non-zero value for the diameter."
                        Else
                    ' Valid numeric input
                    cancelled = False
                    Exit Do
                End If
            Else
                MsgBox "Invalid input. Please enter a numeric value."
            End If
            Loop Until hdiam > 0 Or cancelled
    
            Case "S", "R"
                ' Square or Rectangle profile, prompt for hlength and hwidth
                Do
                hlengthInput = InputBox("Please enter the length for the square/rectangle profile.", "FC Racing 2024 " & Chr(169))
                If hlengthInput = "" Then
                    cancelled = True
                    MsgBox "Input cancelled. Exiting."
                    Exit Sub ' Exit the function if the user cancelled
                ElseIf IsNumeric(hlengthInput) Then
                    hlength = CSng(hlengthInput)
                        If hlength < 0 Then
                            cancelled = False
                            MsgBox "Please enter a positive value for the length."
                        ElseIf hlength = 0 Then
                            cancelled = False
                            MsgBox "Please enter a non-zero value for the length."
                        Else
                    ' Valid numeric input
                    cancelled = False
                    Exit Do
                End If
            Else
                MsgBox "Invalid input. Please enter a numeric value."
            End If
            Loop Until hlength > 0 Or cancelled
    
                Do
                hwidthInput = InputBox("Please enter the width for the square/rectangle profile.", "FC Racing 2024 " & Chr(169))
                If hwidthInput = "" Then
                    cancelled = True
                    MsgBox "Input cancelled. Exiting."
                    Exit Sub ' Exit the function if the user cancelled
                ElseIf IsNumeric(hwidthInput) Then
                    hwidth = CSng(hwidthInput)
                        If hwidth < 0 Then
                            cancelled = False
                            MsgBox "Please enter a positive value for the length."
                        ElseIf hwidth = 0 Then
                            cancelled = False
                            MsgBox "Please enter a non-zero value for the length."
                        Else
                    ' Valid numeric input
                    cancelled = False
                    Exit Do
                End If
            Else
                MsgBox "Invalid input. Please enter a numeric value."
            End If
            Loop Until hwidth > 0 Or cancelled
        End Select
        
        ' Validate and get centre distance
        If Not SetParameters("centre distance", cdist) Then Exit Sub
        
        ' Validate and get quantity
        Do
            quantityInput = InputBox("Please enter the quantity (must be a positive integer):", "FC Racing 2024 " & Chr(169))
        
            If quantityInput = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
                Exit Sub ' Exit the function if the user cancelled
            ElseIf IsNumeric(quantityInput) Then
                If quantityInput <= 0 Or quantityInput <> Int(quantityInput) Then
                    MsgBox "Please enter a positive integer value for the quantity."
                Else
                    ' Valid numeric input
                    cancelled = False
                    Exit Do
                End If
            Else
                MsgBox "Invalid input. Please enter a numeric value."
            End If
        Loop Until cancelled
        quantity = Val(quantityInput)
        
        ' Validate and get punch type
        Do
            punchInput = InputBox("Please enter 1 for single punched holes or 2 for double punched holes.", "FC Racing 2024 " & Chr(169))
            If punchInput = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
            Exit Sub ' Exit the function if the user canceled
            ElseIf punchInput <> "1" And punchInput <> "2" Then
                MsgBox "Please enter either the number 1 or 2 only."
            End If
        Loop Until punchInput = "1" Or punchInput = "2" Or cancelled
    
        ' Convert the input to an integer
        punch = Val(punchInput)
        
        ' Determine hole side, if sheight = swidth then hsInput=0
        If sheight = swidth Then
            hsInput = 0
        Else
            Do
            hsInput = InputBox("Please enter " & sheight & " or " & swidth & " to confirm which face to cut.", "FC Racing 2024 " & Chr(169))
        
                If hsInput = "" Then
                    cancelled = True
                    MsgBox "Input cancelled. Exiting."
                Exit Sub ' Exit the function if the user canceled
                ElseIf hsInput <> sheight And hsInput <> swidth Then
                    MsgBox "Please enter either the number " & sheight & " or " & swidth & " only."
                End If
            Loop Until hsInput = sheight Or hsInput = swidth Or cancelled
        End If
    
        ' custom cut length
        clengthnyc = MsgBox("Would you like to specify a custom cutting length?", vbYesNoCancel, "FC Racing 2024 " & Chr(169))
        
        If clengthnyc = 6 Then
        Do
            clengthinput = InputBox("Please specify the desired cutting length: ", "FC Racing 2024 " & Chr(169))
            If clengthinput = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
                Exit Sub ' Exit the function if the user cancelled
            ElseIf IsNumeric(clengthinput) Then
                clength = CSng(clengthinput)
                If clength <= 0 Then
                    MsgBox "Please enter a positive value for the cutting length."
                Else
                    ' Valid numeric input
                    cancelled = False
                    Exit Do
                End If
            Else
                MsgBox "Invalid input. Please enter a numeric value."
            End If
        Loop Until cancelled
        
        ElseIf clengthnyc = 7 Then
            fclength = flength
        
        Else
            cancelled = True
            MsgBox "Input cancelled. Exiting."
            Exit Sub
        End If
        
        ' rake calculations
        rakenyc = MsgBox("Would you like to specify a custom rake between 0 and 45 degrees?", vbYesNoCancel, "FC Racing 2024 " & Chr(169))
        
        If rakenyc = 6 Then
        Do
            rakeInput = InputBox("Please enter the desired rake in degrees:", "FC Racing 2024 " & Chr(169))
        
            If rakeInput = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
                Exit Sub ' Exit the function if the user cancelled
            ElseIf IsNumeric(rakeInput) Then
                If rakeInput <= 0 Or rakeInput >= 45 Then
                    MsgBox "Please enter a positive value between 0 and 45 degrees for the rake."
                Else
                    ' Valid numeric input
                    cancelled = False
                    Exit Do
                End If
            Else
                MsgBox "Invalid input. Please enter a numeric value."
            End If
        Loop Until cancelled
        
        rake = (rakeInput * 3.14159265359) / 180
            If hsInput = sheight Then
                rlength = (length * Cos(rake)) - ((swidth) * Sin(rake))
            ElseIf hsInput = swidth Then
                rlength = (length * Cos(rake)) - ((sheight) * Sin(rake))
            Else
                rlength = (length * Cos(rake)) - ((swidth) * Sin(rake))
            End If
            
            flength = rlength
            
                If clengthnyc = 7 Then
                    fclength = rlength
                ElseIf clengthnyc = 6 Then
                    If hsInput = sheight Then
                        fclength = (clength * Cos(rake)) - ((swidth) * Sin(rake))
                    ElseIf hsInput = swidth Then
                        fclength = (clength * Cos(rake)) - ((sheight) * Sin(rake))
                    Else
                        fclength = (clength * Cos(rake)) - ((swidth) * Sin(rake))
                    End If
                Else
                    fclength = rlength
                End If
        
        ElseIf rakenyc = 7 Then
            rakeInput = 0
            rake = 0
            flength = length
                If clengthnyc = 7 Then
                    fclength = flength
                ElseIf clengthnyc = 6 Then
                    fclength = clength
                Else
                End If
            
        Else
            cancelled = True
            MsgBox "Input cancelled. Exiting."
            Exit Sub
        End If
    
        ' hole number raw calc
        hnumraw = Int(fclength / cdist)
        
        ' gap size calc
        If hprofile Like "C" Then
            gsize = cdist - hdiam
        Else
            gsize = cdist - hlength
        End If
        
        ' edge gap raw based on int(clength/cdist)
        If hprofile Like "C" Then
            egapraw = (fclength - (hnumraw) * (hdiam) - (hnumraw - 1) * (gsize)) / 2
        Else
            egapraw = (fclength - (hnumraw) * (hlength) - (hnumraw - 1) * (gsize)) / 2
        End If
        
        
        ' determining hnum from hnum raw calc depending on hprofile
        If hprofile Like "C" Then
            If egapraw < 40 Then
                hnum = hnumraw - 1
                egap40 = (fclength - (hnum) * (hdiam) - (hnum - 1) * (gsize)) / 2
                
            ElseIf egapraw > (gsize + 2) Then
                hnum = hnumraw + 1
                egap2 = (fclength - (hnum) * (hdiam) - (hnum - 1) * (gsize)) / 2
    
            Else
                hnum = hnumraw
                'egapF = egapraw
            End If
        
        Else
            If egapraw < 40 Then
                hnum = hnumraw - 1
                egap40 = (fclength - (hnum) * (hlength) - (hnum - 1) * (gsize)) / 2
                
            ElseIf egapraw > (gsize + 2) Then
                hnum = hnumraw + 1
                egap2 = (fclength - (hnum) * (hlength) - (hnum - 1) * (gsize)) / 2
            
            Else
                hnum = hnumraw
                'egapF = egapraw
            End If
        End If
        
        ' converting hnum to hnumF through finalising ambiguities
        If (egapraw < 40) And (egap40 > gsize + 2) Then
        Do
            hnumFinput = InputBox("Please enter either: " & vbCrLf _
            & hnumraw & " for edge gap = " & egapraw & vbCrLf & "-OR-" & vbCrLf _
            & hnumraw - 1 & " for edge gap = " & egap40 & vbCrLf _
            & "For a universal gap size = " & gsize & vbCrLf & vbCrLf & "! IMPORTANT !" & vbCrLf _
            & "For full length pieces, please choose " & hnumraw - 1 & " to allow enough space to punch holes. " _
            & "Please note that the left edge gap will always default to universal gap size for full length pieces." & vbCrLf _
            & vbCrLf & "If you wish to define a custom edge gap, please do so in the custom gap size and centering prompt.", "FC Racing 2024 " & Chr(169))
            If hnumFinput = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
                Exit Sub ' Exit the function if the user canceled
            ElseIf IsNumeric(hnumFinput) Then
                If (CInt(hnumFinput) = hnumraw) Or (CInt(hnumFinput) = hnumraw - 1) Then
                    Exit Do ' Exit the loop if the input is valid
                Else
                    MsgBox "Please enter either the number " & hnumraw & " or " & hnumraw - 1 & " only."
                End If
            Else
                MsgBox "Please enter a numeric value corresponding to " & hnumraw & " or " & hnumraw - 1 & " only."
            End If
        Loop Until cancelled
            hnumF = CInt(hnumFinput)
    
        ElseIf (egapraw > gsize + 2) And (egap2 < 40) Then
        Do
            hnumFinput = InputBox("Please enter either: " & vbCrLf _
            & hnumraw & " for edge gap = " & egapraw & vbCrLf & "-OR-" & vbCrLf _
            & hnumraw + 1 & " for edge gap = " & egap2 & vbCrLf _
            & "For a universal gap size = " & gsize & vbCrLf & vbCrLf & "! IMPORTANT !" & vbCrLf _
            & "For full length pieces, please choose " & hnumraw & " to allow enough space to punch holes. " _
            & "Please note that the left edge gap will always default to universal gap size for full length pieces." & vbCrLf _
            & vbCrLf & "If you wish to define a custom edge gap, please do so in the custom gap size and centering prompt.", "FC Racing 2024 " & Chr(169))
            If hnumFinput = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
                Exit Sub ' Exit the function if the user canceled
            ElseIf IsNumeric(hnumFinput) Then
                If (CInt(hnumFinput) = hnumraw) Or (CInt(hnumFinput) = hnumraw + 1) Then
                    Exit Do ' Exit the loop if the input is valid
                Else
                    MsgBox "Please enter either the number " & hnumraw & " or " & hnumraw + 1 & " only."
                End If
            Else
                MsgBox "Please enter a numeric value corresponding to " & hnumraw & " or " & hnumraw + 1 & " only."
            End If
        Loop Until cancelled
            hnumF = CInt(hnumFinput)
    
        Else
            hnumF = CInt(hnumraw)
        End If
        
    
        ' custom edge gap or no custom edge gap + final edge gap calcs based on centering
        cgap = MsgBox("Would you like to specify a custom starting gap?", vbYesNoCancel, "FC Racing 2024 " & Chr(169))
        If fclength <> flength Then
            If cgap = 6 Then
                    Do
                    egleftinput = InputBox("Please specify the desired starting gap: ", "FC Racing 2024 " & Chr(169))
                    If egleftinput = "" Then
                        cancelled = True
                        MsgBox "Input cancelled. Exiting."
                        Exit Sub ' Exit the function if the user cancelled
                    ElseIf IsNumeric(egleftinput) Then
                        egleft = CSng(egleftinput)
                            If egleft < 0 Then
                               cancelled = False
                                MsgBox "Please enter a positive value for the starting gap."
                            ElseIf egleft = 0 Then
                                cancelled = False
                                MsgBox "Please enter a non-zero value for the starting gap."
                            Else
                        ' Valid numeric input
                        cancelled = False
                        Exit Do
                        End If
                    Else
                        MsgBox "Invalid input. Please enter a numeric value."
                    End If
                    Loop Until egleft > 0 Or cancelled
                
                If hsInput = sheight Then
                    If hprofile Like "C" Then
                        egright = fclength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize) - egleft + (flength - fclength)
                    Else
                        egright = fclength - (hnumF) * (hlength) - (hnumF - 1) * (gsize) - egleft + (flength - fclength)
                    End If
                Else
                    If hprofile Like "C" Then
                        egright = fclength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize) - egleft + (flength - fclength)
                    Else
                        egright = fclength - (hnumF) * (hlength) - (hnumF - 1) * (gsize) - egleft + (flength - fclength)
                    End If
                End If
            
                ElseIf cgap = 7 Then
                    If (matInput Like "A" And length >= 5800) Or ((matInput Like "G" Or matInput Like "S") And length >= 7800) Then
                        egleft = gsize
            
                        If hsInput = sheight Then
                            If hprofile Like "C" Then
                                egright = fclength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize) - egleft + (flength - fclength)
                            Else
                                egright = fclength - (hnumF) * (hlength) - (hnumF - 1) * (gsize) - egleft + (flength - fclength)
                            End If
                        Else
                            If hprofile Like "C" Then
                                egright = fclength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize) - egleft + (flength - fclength)
                            Else
                                egright = fclength - (hnumF) * (hlength) - (hnumF - 1) * (gsize) - egleft + (flength - fclength)
                            End If
                        End If
                    Else
                            If hprofile Like "C" Then
                                egleft = (flength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize)) / 2
                            Else
                                egleft = (flength - (hnumF) * (hlength) - (hnumF - 1) * (gsize)) / 2
                            End If
                        If hsInput = sheight Then
                            egright = egleft + (flength - fclength)
                        Else
                            egright = egleft + (flength - fclength)
                        End If
                    End If
            
                Else
                    cancelled = True
                    MsgBox "Input cancelled. Exiting."
                    Exit Sub
            End If
        Else
        If cgap = 6 Then
                    Do
                    egleftinput = InputBox("Please specify the desired starting gap: ", "FC Racing 2024 " & Chr(169))
                    If egleftinput = "" Then
                        cancelled = True
                        MsgBox "Input cancelled. Exiting."
                        Exit Sub ' Exit the function if the user cancelled
                    ElseIf IsNumeric(egleftinput) Then
                        egleft = CSng(egleftinput)
                            If egleft < 0 Then
                               cancelled = False
                                MsgBox "Please enter a positive value for the starting gap."
                            ElseIf egleft = 0 Then
                                cancelled = False
                                MsgBox "Please enter a non-zero value for the starting gap."
                            Else
                        ' Valid numeric input
                        cancelled = False
                        Exit Do
                        End If
                        
                    Else
                        MsgBox "Invalid input. Please enter a numeric value."
                    End If
                    Loop Until egleft > 0 Or cancelled
                If hsInput = sheight Then
                    If hprofile Like "C" Then
                        egright = flength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize) - egleft
                    Else
                        egright = flength - (hnumF) * (hlength) - (hnumF - 1) * (gsize) - egleft
                    End If
                Else
                    If hprofile Like "C" Then
                        egright = flength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize) - egleft
                    Else
                        egright = flength - (hnumF) * (hlength) - (hnumF - 1) * (gsize) - egleft
                    End If
                End If
                                
                ElseIf cgap = 7 Then
                    If (matInput Like "A" And length >= 6000) Or ((matInput Like "G" Or matInput Like "S") And length >= 7800) Then
                        egleft = gsize
                        If hsInput = sheight Then
                            If hprofile Like "C" Then
                                egright = flength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize) - egleft
                            Else
                                egright = flength - (hnumF) * (hlength) - (hnumF - 1) * (gsize) - egleft
                            End If
                        Else
                            If hprofile Like "C" Then
                                egright = flength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize) - egleft
                            Else
                                egright = flength - (hnumF) * (hlength) - (hnumF - 1) * (gsize) - egleft
                            End If
                        End If
                    Else
                            If hprofile Like "C" Then
                                egleft = (flength - (hnumF) * (hdiam) - (hnumF - 1) * (gsize)) / 2
                            Else
                                egleft = (flength - (hnumF) * (hlength) - (hnumF - 1) * (gsize)) / 2
                            End If
                        If hsInput = sheight Then
                            egright = egleft
                        Else
                            egright = egleft
                        End If
                    End If
            
                Else
                    cancelled = True
                    MsgBox "Input cancelled. Exiting."
                    Exit Sub
            End If
        End If
        
        ' Confirm order number
        Do
            sordernum = InputBox("Please enter a customer sales order number or part quotation number:", "FC Racing 2024 " & Chr(169))
            If sordernum = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
            Exit Sub ' Exit the function if the user canceled
            Else
            End If
        Loop Until sordernum <> ""
        
        ' Confirm customer name
        Do
            cordername = InputBox("Please enter the customer name:", "FC Racing 2024 " & Chr(169))
            If cordername = "" Then
                cancelled = True
                MsgBox "Input cancelled. Exiting."
            Exit Sub ' Exit the function if the user canceled
            Else
            
        End If
        Loop Until cordername <> ""
    
        
        ' Display a message box with OK and Cancel buttons
        Dim result As VbMsgBoxResult
        result = MsgBox("Final data before construction: " & vbCrLf & "Gap Size = " & gsize & vbCrLf & "Final Hole Number = " _
        & hnumF & vbCrLf & "Left Edge Gap = " & egleft & vbCrLf & "Right Edge Gap = " & egright _
        & vbCrLf & "Length = " & length & vbCrLf & "Rake in degrees (if applicable) = " & rakeInput _
        & vbCrLf & "Raked Length (if applicable) = " & rlength & vbCrLf & "Length to Cut = " & fclength _
        & vbCrLf & vbCrLf & "Please press OK to continue", vbOKCancel + vbInformation, "FC Racing 2024 " & Chr(169))
        
        ' Check the result
        If result = vbCancel Then
            ' User clicked Cancel, perform cancellation actions
            MsgBox "Input cancelled. Exiting.", vbOKOnly, "FC Racing 2024 " & Chr(169)
            Exit Sub ' Exit the subroutine or perform other cancellation actions
        End If
    
        'make a new rhs or shs with correct dimensions
        'create a new part doc
        'draw a line on the front plane
        'dimension the line on front plane to length
        'go to weldments and choose structural member on line
        'dimension weldment to sheight, swidth and thickness
            'if GAL, no change
            'if ALU, based on thickness set radius
            
        'Dim swApp As Object ' SolidWorks.SldWorks
        Dim partDoc As Object ' SolidWorks.ModelDoc2
        'Dim swModel As Object ' SolidWorks.ModelDoc2
        Dim swSketchMgr As Object ' SolidWorks.SketchManager
        Dim swSketch As Object ' SolidWorks.Sketch
        Dim swSketchSeg As Object
        Dim swDim As Object
        Dim boolstatus As Boolean
    
        ' Create or get an instance of SolidWorks
        On Error Resume Next
        Set swApp = GetObject(, "SldWorks.Application")
        On Error GoTo 0
    
        If swApp Is Nothing Then
            Set swApp = CreateObject("SldWorks.Application")
            swApp.Visible = True ' Show SolidWorks
        End If
    
        ' Create a new part document
        Set partDoc = swApp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2022\templates\Part.PRTDOT", 0, 0, 0)
        
        ' Get the active document
        Set swModel = swApp.ActiveDoc
        
            ' create line1
            boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            swModel.SketchManager.InsertSketch True
            swModel.ClearSelection2 True
            Dim skSegment As Object
            Set skSegment = swModel.SketchManager.CreateLine(0#, 0#, 0#, 0.1, 0#, 0#)
            swModel.ClearSelection2 True
            
            ' Select Line1
            swModel.Extension.SelectByID2 "Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0
            
            ' Smart dimension: length
            Dim myDisplayDim As Object
            SendKeys "~"
            Set myDisplayDim = swModel.AddDimension2(0.03, 0.015, 0)
            SendKeys "~"
            swModel.ClearSelection2 True
            Dim myDimension As Object
            Set myDimension = swModel.parameter("D1@Sketch1")
            myDimension.SystemValue = length / 1000
            swModel.ClearSelection2 True
            swModel.SketchManager.InsertSketch True
    
            ' Select Line1
            swModel.Extension.SelectByID2 "Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
            
            ' make rhs or shs
            Dim myFeature As Object
            Dim vGroups As Variant
            Dim GroupArray() As Object
            ReDim GroupArray(0 To 0) As Object
            Dim Group1 As Object
            Set Group1 = swModel.FeatureManager.CreateStructuralMemberGroup()
            Dim vSegement1 As Variant
            Dim SegementArray1() As Object
            ReDim SegementArray1(0 To 0) As Object
            swModel.ClearSelection2 True
            boolstatus = swModel.Extension.SelectByID2("Line1@Sketch1", "EXTSKETCHSEGMENT", -5.01299348120124E-02, 4.61760461760462E-02, 0, True, 0, Nothing, 0)
            Dim Segment As Object
            Set Segment = swModel.SelectionManager.GetSelectedObject5(1)
            Set SegementArray1(0) = Segment
            vSegement1 = SegementArray1
            Group1.Segments = (vSegement1)
            Group1.ApplyCornerTreatment = True
            Group1.CornerTreatmentType = 1
            Group1.GapWithinGroup = 0
            Group1.GapForOtherGroups = 0
            Group1.Angle = 0
            Set GroupArray(0) = Group1
            vGroups = GroupArray
            Set myFeature = swModel.FeatureManager.InsertStructuralWeldment4("C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\data\weldment profiles\AS\rhs c350 as.sldlfp", 1, True, (vGroups))
            swModel.ClearSelection2 True
            
            ' dimension rhs
            boolstatus = swModel.Extension.SelectByID2("Sketch11", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            swModel.EditSketch
            swModel.ClearSelection2 True
            boolstatus = swModel.Extension.SelectByID2("Width@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, 3.57280691191519E-03, 5.96791080471766E-02, False, 0, Nothing, 0)
            Set myDimension = swModel.parameter("Width@Sketch11")
            myDimension.SystemValue = swidth / 1000
            boolstatus = swModel.Extension.SelectByID2("Depth@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -4.48585756718246E-02, -1.07184207357457E-02, False, 0, Nothing, 0)
            Set myDimension = swModel.parameter("Depth@Sketch11")
            myDimension.SystemValue = sheight / 1000
            boolstatus = swModel.Extension.SelectByID2("Thickness@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.75993821957306E-02, 1.75993821957306E-02, False, 0, Nothing, 0)
            Set myDimension = swModel.parameter("Thickness@Sketch11")
            myDimension.SystemValue = thickness / 1000
            
            ' dimension rhs radius based on ALU or GAL
            If matInput Like "A" And thickness = 1.6 Then
                boolstatus = swModel.Extension.SelectByID2("D2@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.33172591687161E-02, 3.04148114989708E-02, False, 0, Nothing, 0)
                Set myDimension = swModel.parameter("D2@Sketch11")
                myDimension.SystemValue = 1 / 1000
                
            ElseIf matInput Like "A" And thickness = 2 Then
                boolstatus = swModel.Extension.SelectByID2("D2@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.33172591687161E-02, 3.04148114989708E-02, False, 0, Nothing, 0)
                Set myDimension = swModel.parameter("D2@Sketch11")
                myDimension.SystemValue = 1.5 / 1000
                
            ElseIf matInput Like "A" And thickness = 2.5 Then
                boolstatus = swModel.Extension.SelectByID2("D2@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.33172591687161E-02, 3.04148114989708E-02, False, 0, Nothing, 0)
                Set myDimension = swModel.parameter("D2@Sketch11")
                myDimension.SystemValue = 2 / 1000
                
                
            ElseIf matInput Like "A" And thickness = 3 Then
                boolstatus = swModel.Extension.SelectByID2("D2@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.33172591687161E-02, 3.04148114989708E-02, False, 0, Nothing, 0)
                Set myDimension = swModel.parameter("D2@Sketch11")
                myDimension.SystemValue = 2.5 / 1000
                
            ElseIf matInput Like "A" And thickness = 3.5 Then
                boolstatus = swModel.Extension.SelectByID2("D2@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.33172591687161E-02, 3.04148114989708E-02, False, 0, Nothing, 0)
                Set myDimension = swModel.parameter("D2@Sketch11")
                myDimension.SystemValue = 3 / 1000
                
    
            ElseIf matInput Like "A" And thickness = 4 Then
                boolstatus = swModel.Extension.SelectByID2("D2@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.33172591687161E-02, 3.04148114989708E-02, False, 0, Nothing, 0)
                Set myDimension = swModel.parameter("D2@Sketch11")
                myDimension.SystemValue = 3.5 / 1000
    
            ElseIf (matInput Like "G" Or matInput Like "P" Or matInput Like "S") And thickness <= 3 Then
                boolstatus = swModel.Extension.SelectByID2("D2@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.33172591687161E-02, 3.04148114989708E-02, False, 0, Nothing, 0)
                Set myDimension = swModel.parameter("D2@Sketch11")
                myDimension.SystemValue = 2 * (thickness / 1000)
            
            Else 'matInput Like "G" or "P" or "S" And thickness > 3
                boolstatus = swModel.Extension.SelectByID2("D2@Sketch11@FC Racing laser template v1 29012024.PRTDOT", "DIMENSION", 3.15, -1.33172591687161E-02, 3.04148114989708E-02, False, 0, Nothing, 0)
                Set myDimension = swModel.parameter("D2@Sketch11")
                myDimension.SystemValue = 2.5 * (thickness / 1000)
    
            End If
            swModel.ClearSelection2 True
            swModel.ClearSelection2 True
            swModel.SketchManager.InsertSketch True
            
            
            ' insert sketch onto specified sheight or swidth face (ADDING RAKE)
            Dim myRefPlane As Object
            Dim pi As Double
            If rakenyc = 7 Then
                If hsInput = swidth Then
                    boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    Set myRefPlane = swModel.FeatureManager.InsertRefPlane(8, (sheight / 1000) / 2, 0, 0, 0, 0)
                    swModel.ClearSelection2 True
                    boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    swModel.ClearSelection2 True
                    boolstatus = swModel.Extension.SelectByID2("Plane2", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    swModel.SketchManager.InsertSketch True
                ElseIf hsInput = sheight Then
                    boolstatus = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    Set myRefPlane = swModel.FeatureManager.InsertRefPlane(8, (swidth / 1000) / 2, 0, 0, 0, 0)
                    swModel.ClearSelection2 True
                    boolstatus = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    swModel.ClearSelection2 True
                    boolstatus = swModel.Extension.SelectByID2("Plane2", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    swModel.SketchManager.InsertSketch True
                Else
                    boolstatus = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                    Set myRefPlane = swModel.FeatureManager.InsertRefPlane(8, (sheight / 1000) / 2, 0, 0, 0, 0)
                    swModel.ClearSelection2 True
                    boolstatus = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    swModel.ClearSelection2 True
                    boolstatus = swModel.Extension.SelectByID2("Plane2", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
                    swModel.SketchManager.InsertSketch True
                End If
                
            Else
                If hsInput = swidth Then
                        boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                        boolstatus = swModel.Extension.SelectByID2("Line3@Sketch11", "EXTSKETCHSEGMENT", 0, 0, ((sheight / 1000) / 2), True, 1, Nothing, 0)
                        Set myRefPlane = swModel.FeatureManager.InsertRefPlane(16, ((180 - rakeInput) * 3.1415926535) / 180, 4, 0, 0, 0)
                        swModel.ClearSelection2 True
                    ElseIf hsInput = sheight Then
                        boolstatus = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                        boolstatus = swModel.Extension.SelectByID2("Line4@Sketch11", "EXTSKETCHSEGMENT", 0, -((swidth / 1000) / 2), 0, True, 1, Nothing, 0)
                        Set myRefPlane = swModel.FeatureManager.InsertRefPlane(16, (rakeInput * 3.1415926535) / 180, 4, 0, 0, 0)
                        swModel.ClearSelection2 True
                    Else
                        boolstatus = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                        boolstatus = swModel.Extension.SelectByID2("Line4@Sketch11", "EXTSKETCHSEGMENT", 0, -((swidth / 1000) / 2), 0, True, 1, Nothing, 0)
                        Set myRefPlane = swModel.FeatureManager.InsertRefPlane(16, (rakeInput * 3.1415926535) / 180, 4, 0, 0, 0)
                        swModel.ClearSelection2 True
                
                End If
            End If
        
            Dim swModelView As Object
            Dim swTranslation() As Double
            Dim swTranslationVar As Variant
            Dim swMathUtils As Object
            Dim swTranslationVector As MathVector
            Set swModelView = swModel.ActiveView
        
        If rakenyc = 6 Then
            ' changing view
            If hsInput = swidth Then
                swModel.ViewRotYPlusNinety
                swModel.ViewRotYPlusNinety
                swModel.ViewRotateminusz
                swModel.ViewRotateminusz
                swModel.ViewRotateminusz
                swModel.ViewRotateminusz
                swModel.ViewRotateminusz
                swModel.ViewRotateminusz
            Else
                swModel.ViewRotXPlusNinety
                swModel.ViewRotXPlusNinety
            End If
            
            ' create coordinate and new plane system
            If hsInput = sheight Or hsInput = 0 Then
                boolstatus = swModel.Extension.SelectByID2("Plane2", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                boolstatus = swModel.InsertCoordinateSystem(False, False, False)
                swModel.ClearSelection2 True
                boolstatus = swModel.Extension.SelectByID2("Plane2", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("Coordinate System1", "COORDSYS", 0, 0, 0, True, 1, Nothing, 0)
                Set myRefPlane = swModel.FeatureManager.InsertRefPlane(1, 0, 4, 0, 0, 0)
                boolstatus = swModel.Extension.SelectByID2("Plane3", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                Set myRefPlane = swModel.FeatureManager.InsertRefPlane(8, (swidth / 2000) * Cos(rake), 0, 0, 0, 0)
                boolstatus = swModel.Extension.SelectByID2("Plane4", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                swModel.SketchManager.InsertSketch True
            ElseIf hsInput = swidth Then
                boolstatus = swModel.Extension.SelectByID2("Plane2", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                boolstatus = swModel.InsertCoordinateSystem(False, False, False)
                swModel.ClearSelection2 True
                boolstatus = swModel.Extension.SelectByID2("Plane2", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("Coordinate System1", "COORDSYS", 0, 0, 0, True, 1, Nothing, 0)
                Set myRefPlane = swModel.FeatureManager.InsertRefPlane(1, 0, 4, 0, 0, 0)
                boolstatus = swModel.Extension.SelectByID2("Plane3", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                Set myRefPlane = swModel.FeatureManager.InsertRefPlane(264, (sheight / 2000) * Cos(rake), 0, 0, 0, 0)
                boolstatus = swModel.Extension.SelectByID2("Plane4", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                
                swModel.Extension.RunCommand 3113, 0
                
                boolstatus = swModel.Extension.SelectByID2("Plane4", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
                swModel.SketchManager.InsertSketch True
                End If
        
                 Set swModelView = swModel.ActiveView
                 swModelView.Scale2 = 2
                 ReDim swTranslation(0 To 2) As Double
                 swTranslation(0) = -2.72113276432473E-02
                 swTranslation(1) = 1.68434983243196E-03
                 swTranslation(2) = 0
                 swTranslationVar = swTranslation
                 Set swMathUtils = swApp.GetMathUtility()
                 Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
                 swModelView.Translation3 = swTranslationVector
                 swModel.ClearSelection2 True
            Else
            End If
        
    ' draw circle/square/rectangle on sketch
        If hprofile Like "C" Then
            If hsInput = sheight Then
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
                ReDim swTranslation(0 To 2) As Double
                swTranslation(0) = -2.72113276432473E-02
                swTranslation(1) = 1.68434983243196E-03
                swTranslation(2) = 0
                swTranslationVar = swTranslation
                Set swMathUtils = swApp.GetMathUtility()
                Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
                swModelView.Translation3 = swTranslationVector
                swModel.ClearSelection2 True
                Set skSegment = swModel.SketchManager.CreateCenterLine(0#, 0#, 0#, 0.1, 0#, 0#)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstToRectEntity, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                Set skSegment = swModel.SketchManager.CreateCircle(0.097855, 0#, 0#, 0.102164, -0.05, 0#)
                swModel.ClearSelection2 True
                boolstatus = swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                SendKeys "~"
                Set myDisplayDim = swModel.AddDimension2(9.70818408942951E-02, 0.01, -3.24709520029243E-02)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D1@Sketch12")
                myDimension.SystemValue = hdiam / 1000
                boolstatus = swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 8.74690978031999E-02, 1.00000000000001E-02, 1.01786897575671E-02, False, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("Line4@Sketch11", "EXTSKETCHSEGMENT", 9.99999977648259E-03, 6.64296815735098E-03, 0, True, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(5.07859862371815E-02, 0.01, 3.66966017330021E-02)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D2@Sketch12")
                myDimension.SystemValue = (egleft + (hdiam / 2)) / 1000
                swModelView.Scale2 = 2
            ElseIf hsInput = swidth Then
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
                ReDim swTranslation(0 To 2) As Double
                swTranslation(0) = -2.72113276432473E-02
                swTranslation(1) = 1.68434983243196E-03
                swTranslation(2) = 0
                swTranslationVar = swTranslation
                Set swMathUtils = swApp.GetMathUtility()
                Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
                swModelView.Translation3 = swTranslationVector
                swModel.ClearSelection2 True
                Set skSegment = swModel.SketchManager.CreateCenterLine(0#, 0#, 0#, 0.1, 0#, 0#)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstToRectEntity, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                Set skSegment = swModel.SketchManager.CreateCircle(0.098629, 0#, 0#, 0.101722, -0.05, 0#)
                swModel.ClearSelection2 True
                boolstatus = swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                SendKeys "~"
                Set myDisplayDim = swModel.AddDimension2(9.84077364930669E-02, 1.82175743161279E-02, 0.025)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D1@Sketch12")
                myDimension.SystemValue = hdiam / 1000
                boolstatus = swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("Line3@Sketch11", "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(5.37692513344179E-02, -1.91284850492764E-02, 0.025)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D2@Sketch12")
                myDimension.SystemValue = (egleft + (hdiam / 2)) / 1000
                swModelView.Scale2 = 2
            Else
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
                ReDim swTranslation(0 To 2) As Double
                swTranslation(0) = -2.72113276432473E-02
                swTranslation(1) = 1.68434983243196E-03
                swTranslation(2) = 0
                swTranslationVar = swTranslation
                Set swMathUtils = swApp.GetMathUtility()
                Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
                swModelView.Translation3 = swTranslationVector
                swModel.ClearSelection2 True
                Set skSegment = swModel.SketchManager.CreateCenterLine(0#, 0#, 0#, 0.1, 0#, 0#)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstToRectEntity, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                Set skSegment = swModel.SketchManager.CreateCircle(0.097855, 0#, 0#, 0.102164, -0.05, 0#)
                swModel.ClearSelection2 True
                boolstatus = swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                SendKeys "~"
                Set myDisplayDim = swModel.AddDimension2(9.70818408942951E-02, 0.01, -3.24709520029243E-02)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D1@Sketch12")
                myDimension.SystemValue = hdiam / 1000
                boolstatus = swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 8.74690978031999E-02, 1.00000000000001E-02, 1.01786897575671E-02, False, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("Line4@Sketch11", "EXTSKETCHSEGMENT", 9.99999977648259E-03, 6.64296815735098E-03, 0, True, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(5.07859862371815E-02, 0.01, 3.66966017330021E-02)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D2@Sketch12")
                myDimension.SystemValue = (egleft + (hdiam / 2)) / 1000
                swModelView.Scale2 = 2
            End If
        
        Else ' hprofile like "s" or "r"
            If hsInput = sheight Then
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
                ReDim swTranslation(0 To 2) As Double
                swTranslation(0) = -2.72113276432473E-02
                swTranslation(1) = 1.68434983243196E-03
                swTranslation(2) = 0
                swTranslationVar = swTranslation
                Set swMathUtils = swApp.GetMathUtility()
                Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
                swModelView.Translation3 = swTranslationVector
                swModel.ClearSelection2 True
                Set skSegment = swModel.SketchManager.CreateCenterLine(0#, 0#, 0#, 0.1, 0#, 0#)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstToRectEntity, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                Dim vSkLines As Variant
                vSkLines = swModel.SketchManager.CreateCenterRectangle(9.76342973937833E-02, 0, 0, 0.107689005684469, -0.1, 0)
                swModel.ClearSelection2 True
                boolstatus = swModel.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                SendKeys "~"
                Set myDisplayDim = swModel.AddDimension2(2.84867495157248E-02, 0.01, -5.24159166164636E-03)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D1@Sketch12")
                myDimension.SystemValue = hlength / 1000
                boolstatus = swModel.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(4.24086533028282E-02, 0.01, 1.00116678257063E-03)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D2@Sketch12")
                myDimension.SystemValue = hwidth / 1000
                boolstatus = swModel.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("Line4@Sketch11", "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(8.10110468460917E-03, 0.01, 3.54246668021648E-03)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D3@Sketch12")
                myDimension.SystemValue = egleft / 1000
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
            ElseIf hsInput = swidth Then
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
                ReDim swTranslation(0 To 2) As Double
                swTranslation(0) = -2.72113276432473E-02
                swTranslation(1) = 1.68434983243196E-03
                swTranslation(2) = 0
                swTranslationVar = swTranslation
                Set swMathUtils = swApp.GetMathUtility()
                Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
                swModelView.Translation3 = swTranslationVector
                swModel.ClearSelection2 True
                Set skSegment = swModel.SketchManager.CreateCenterLine(0#, 0#, 0#, 0.1, 0#, 0#)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstToRectEntity, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                vSkLines = swModel.SketchManager.CreateCenterRectangle(9.76342973937833E-02, 0, 0, 0.107689005684469, -0.1, 0)
                swModel.ClearSelection2 True
                boolstatus = swModel.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                SendKeys "~"
                Set myDisplayDim = swModel.AddDimension2(2.00341650735549E-02, 6.07027641087871E-03, 0.025)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D1@Sketch12")
                myDimension.SystemValue = hlength / 1000
                boolstatus = swModel.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(2.70503626170554E-02, 1.42964181517758E-03, 0.025)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D2@Sketch12")
                myDimension.SystemValue = hwidth / 1000
                boolstatus = swModel.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("Line3@Sketch11", "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(5.55980478696331E-03, -1.27684902216699E-02, 0.025)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D3@Sketch12")
                myDimension.SystemValue = egleft / 1000
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
            Else
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
                ReDim swTranslation(0 To 2) As Double
                swTranslation(0) = -2.72113276432473E-02
                swTranslation(1) = 1.68434983243196E-03
                swTranslation(2) = 0
                swTranslationVar = swTranslation
                Set swMathUtils = swApp.GetMathUtility()
                Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
                swModelView.Translation3 = swTranslationVector
                swModel.ClearSelection2 True
                Set skSegment = swModel.SketchManager.CreateCenterLine(0#, 0#, 0#, 0.1, 0#, 0#)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstToRectEntity, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                boolstatus = swModel.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
                vSkLines = swModel.SketchManager.CreateCenterRectangle(9.76342973937833E-02, 0, 0, 0.107689005684469, -0.1, 0)
                swModel.ClearSelection2 True
                boolstatus = swModel.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                SendKeys "~"
                Set myDisplayDim = swModel.AddDimension2(2.84867495157248E-02, 0.01, -5.24159166164636E-03)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D1@Sketch12")
                myDimension.SystemValue = hlength / 1000
                boolstatus = swModel.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(4.24086533028282E-02, 0.01, 1.00116678257063E-03)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D2@Sketch12")
                myDimension.SystemValue = hwidth / 1000
                boolstatus = swModel.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("Line4@Sketch11", "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(8.10110468460917E-03, 0.01, 3.54246668021648E-03)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D3@Sketch12")
                myDimension.SystemValue = egleft / 1000
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 2
            End If
        
        ' fillet square/rectangle hole
        If (hlength = 65.8 And hwidth = 16.8) Or (hlength = 38.8 And hwidth = 16.8) Then
            boolstatus = swModel.Extension.SelectByID2("Point8", "SKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Point10", "SKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Point7", "SKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Point9", "SKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
            Set skSegment = swModel.SketchManager.CreateFillet(0.003, 1)
            Set skSegment = swModel.SketchManager.CreateFillet(0.003, 1)
            Set skSegment = swModel.SketchManager.CreateFillet(0.003, 1)
            Set skSegment = swModel.SketchManager.CreateFillet(0.003, 1)
        Else
            boolstatus = swModel.Extension.SelectByID2("Point8", "SKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Point10", "SKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Point7", "SKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Point9", "SKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
            Set skSegment = swModel.SketchManager.CreateFillet(0.001, 1)
            Set skSegment = swModel.SketchManager.CreateFillet(0.001, 1)
            Set skSegment = swModel.SketchManager.CreateFillet(0.001, 1)
            Set skSegment = swModel.SketchManager.CreateFillet(0.001, 1)
        End If
    End If

    
    ' apply linear sketch pattern to hole profiles (accounting for rake)
    If rakenyc = 6 Then
        If hprofile Like "C" Then
            swModel.SketchManager.InsertSketch True
            swModel.SketchManager.InsertSketch True
            Set swModelView = swModel.ActiveView
            swModelView.Scale2 = 0.865195294741085
            ReDim swTranslation(0 To 2) As Double
            swTranslation(0) = -4.64383412938853E-03
            swTranslation(1) = 2.15687003220338E-03
            swTranslation(2) = 0
            swTranslationVar = swTranslation
            Set swMathUtils = swApp.GetMathUtility()
            Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
            swModelView.Translation3 = swTranslationVector
            'swModel.ClearSelection2 True
            'swModel.ViewZoomtofit2
            If hsInput = sheight Or hsInput = 0 Then
                boolstatus = swModel.Extension.SelectByID2("Sketch12", "SKETCHCONTOUR", ((egleft + (hdiam / 2)) / 1000) * Cos(rake), ((egleft + (hdiam / 2)) / 1000) * Sin(rake), 0, False, 0, Nothing, 0)
                boolstatus = swModel.SketchManager.CreateLinearSketchStepAndRepeat(hnumF, 1, (cdist / 1000), 0.01, 0, 1.5707963267949, "", False, False, False, False, False)
            Else
                boolstatus = swModel.Extension.SelectByID2("Sketch12", "SKETCHCONTOUR", ((egleft + (hdiam / 2)) / 1000) * Cos(rake), 0, ((egleft + (hdiam / 2)) / 1000) * Sin(rake), False, 0, Nothing, 0)
                boolstatus = swModel.SketchManager.CreateLinearSketchStepAndRepeat(hnumF, 1, (cdist / 1000), 0.01, 0, 1.5707963267949, "", False, False, False, False, False)
            End If
        Else

                swModel.SketchManager.InsertSketch True
                swModel.SketchManager.InsertSketch True
                Set swModelView = swModel.ActiveView
                swModelView.Scale2 = 0.865195294741085
                ReDim swTranslation(0 To 2) As Double
                swTranslation(0) = -4.64383412938853E-03
                swTranslation(1) = 2.15687003220338E-03
                swTranslation(2) = 0
                swTranslationVar = swTranslation
                Set swMathUtils = swApp.GetMathUtility()
                Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
                swModelView.Translation3 = swTranslationVector


                'swModel.ClearSelection2 True
                'swModel.ViewZoomtofit2
            
            If hsInput = sheight Or hsInput = 0 Then
                boolstatus = swModel.Extension.SelectByID2("Sketch12", "SKETCHCONTOUR", ((egleft + (hlength / 2)) / 1000) * Cos(rake), ((egleft + (hlength / 2)) / 1000) * Sin(rake), 0, False, 0, Nothing, 0)
                boolstatus = swModel.SketchManager.CreateLinearSketchStepAndRepeat(hnumF, 1, (cdist / 1000), 0.01, 0, 1.5707963267949, "", False, False, False, False, False)
            Else
                boolstatus = swModel.Extension.SelectByID2("Sketch12", "SKETCHCONTOUR", ((egleft + (hlength / 2)) / 1000) * Cos(rake), 0, ((egleft + (hlength / 2)) / 1000) * Sin(rake), False, 0, Nothing, 0)
                boolstatus = swModel.SketchManager.CreateLinearSketchStepAndRepeat(hnumF, 1, (cdist / 1000), 0.01, 0, 1.5707963267949, "", False, False, False, False, False)
            End If
        End If
    Else
        If hprofile Like "C" Then
            swModel.SketchManager.InsertSketch True
            swModel.SketchManager.InsertSketch True
            Set swModelView = swModel.ActiveView
            swModelView.Scale2 = 0.865195294741085
            ReDim swTranslation(0 To 2) As Double
            swTranslation(0) = -4.64383412938853E-03
            swTranslation(1) = 2.15687003220338E-03
            swTranslation(2) = 0
            swTranslationVar = swTranslation
            Set swMathUtils = swApp.GetMathUtility()
            Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
            swModelView.Translation3 = swTranslationVector
            swModel.ClearSelection2 True
            boolstatus = swModel.Extension.SelectByID2("Sketch12", "SKETCHCONTOUR", (egleft + (hdiam / 2)) / 1000, 0.0001, 0, False, 0, Nothing, 0)
            boolstatus = swModel.SketchManager.CreateLinearSketchStepAndRepeat(hnumF, 1, (cdist / 1000), 0.01, 0, 1.5707963267949, "", False, False, False, False, False)
        Else
            swModel.SketchManager.InsertSketch True
            swModel.SketchManager.InsertSketch True
            Set swModelView = swModel.ActiveView
            swModelView.Scale2 = 0.865195294741085
            ReDim swTranslation(0 To 2) As Double
            swTranslation(0) = -4.64383412938853E-03
            swTranslation(1) = 2.15687003220338E-03
            swTranslation(2) = 0
            swTranslationVar = swTranslation
            Set swMathUtils = swApp.GetMathUtility()
            Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
            swModelView.Translation3 = swTranslationVector
            swModel.ClearSelection2 True
            boolstatus = swModel.Extension.SelectByID2("Sketch12", "SKETCHCONTOUR", (egleft + (hlength / 2)) / 1000, 0.0001, 0, False, 0, Nothing, 0)
            boolstatus = swModel.SketchManager.CreateLinearSketchStepAndRepeat(hnumF, 1, (cdist / 1000), 0.01, 0, 1.5707963267949, "", False, False, False, False, False)
        End If
        
    End If
    
    ' reset view model
    
    Set swModelView = swModel.ActiveView
    swModelView.Scale2 = 0.327620809267629
    ReDim swTranslation(0 To 2) As Double
    swTranslation(0) = -2.11279515949394
    swTranslation(1) = -3.66131336565298E-03
    swTranslation(2) = -1.63810770781093E-05
    swTranslationVar = swTranslation
    Set swMathUtils = swApp.GetMathUtility()
    Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
    swModelView.Translation3 = swTranslationVector
    
    
    'swModel.ClearSelection2 True
    'swModel.SketchManager.InsertSketch True
    'boolstatus = swModel.Extension.SelectByID2("Sketch12", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
    
    swModel.Extension.SelectByID2 "Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
    swModel.BlankSketch
    
    On Error Resume Next
    
    ' lock edge gap right to linear sketch pattern (currently UNreliable)
    If rakenyc = 6 Then
        If hprofile Like "C" Then
            boolstatus = swModel.Extension.SelectByID2("Arc" & hnumF, "SKETCHSEGMENT", ((flength - egright - (hdiam / 2)) * Cos(rake)) / 1000, _
            ((flength - egright - (hdiam / 2)) * Sin(rake)) / 1000, 0, False, 0, Nothing, 0)
            If hsInput = sheight Or hsInput = 0 Then
                boolstatus = swModel.Extension.SelectByRay(length / 1000, swidth / 2000, 0, 0, -1, 0, 0.001, 1, True, 0, 0)
                Set myDisplayDim = swModel.AddDimension2(fclength / 1000, 0.01, -3.71439881519581E-02)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D4@Sketch12")
                myDimension.SystemValue = ((egright + (hdiam / 2) + swidth * Sin(rake))) / 1000
            Else
                boolstatus = swModel.Extension.SelectByID2("", "EDGE", length / 1000, sheight / 2000, 0, True, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(fclength / 1000, 0.01, -3.71439881519581E-02)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D4@Sketch12")
                myDimension.SystemValue = ((egright + (hdiam / 2) + sheight * Sin(rake))) / 1000
            End If
        Else
            boolstatus = swModel.Extension.SelectByID2("Line" & (hnumF + 1) * 4, "SKETCHSEGMENT", ((flength - egright) * Cos(rake)) / 1000, _
            ((flength - egright) * Sin(rake)) / 1000, 0, False, 0, Nothing, 0)
            If hsInput = sheight Or hsInput = 0 Then
                boolstatus = swModel.Extension.SelectByRay(length / 1000, swidth / 2000, 0, 0, -1, 0, 0.001, 1, True, 0, 0)
                Set myDisplayDim = swModel.AddDimension2(fclength / 1000, 0.01, -3.71439881519581E-02)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D6@Sketch12")
                myDimension.SystemValue = (egright + swidth * Sin(rake)) / 1000
            Else
                boolstatus = swModel.Extension.SelectByID2("", "EDGE", length / 1000, 0, sheight / 2000, True, 0, Nothing, 0)
                Set myDisplayDim = swModel.AddDimension2(fclength / 1000, 0.01, -3.71439881519581E-02)
                SendKeys "~"
                swModel.ClearSelection2 True
                Set myDimension = swModel.parameter("D6@Sketch12")
                myDimension.SystemValue = (egright + sheight * Sin(rake)) / 1000
            End If
        End If
    Else
        If hprofile Like "C" Then
        boolstatus = swModel.Extension.SelectByID2("", "SKETCHCONTOUR", (length - egright - (hdiam / 2)) / 1000, 0, 0, False, 0, Nothing, 0)
            If hsInput = sheight Or hsInput = 0 Then
                boolstatus = swModel.Extension.SelectByRay(length / 1000, 0, 0.001, 0, -1, 0, 0.002, 1, True, 0, 0)
            Else
                boolstatus = swModel.Extension.SelectByID2("", "EDGE", length / 1000, 0.001, 0, True, 0, Nothing, 0)
                'boolstatus = swModel.Extension.SelectByRay(length / 1000, 0.001, 0, 0, -1, 0, 0.002, 1, True, 0, 0)
            End If
        Set myDisplayDim = swModel.AddDimension2(fclength / 1000, 0.01, -3.71439881519581E-02)
        SendKeys "~"
        swModel.ClearSelection2 True
        Set myDimension = swModel.parameter("D4@Sketch12")
        myDimension.SystemValue = (egright + (hdiam / 2)) / 1000
        Else
            boolstatus = swModel.Extension.SelectByID2("Line" & (hnumF + 1) * 4, "SKETCHSEGMENT", (length - egright) / 1000, 0, 0, False, 0, Nothing, 0)
            If hsInput = sheight Or hsInput = 0 Then
                'boolstatus = swModel.Extension.SelectByID2("Line" & (hnumF + 1) * 4, "SKETCHSEGMENT", (length - egright) / 1000, 0, 0, False, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByRay(length / 1000, 0, 0.001, 0, -1, 0, 0.002, 1, True, 0, 0)
            Else
                'boolstatus = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", (length - egright) / 1000, 0, 0, False, 0, Nothing, 0)
                boolstatus = swModel.Extension.SelectByID2("", "EDGE", length / 1000, 0, 0, True, 0, Nothing, 0)
            End If
        Set myDisplayDim = swModel.AddDimension2(fclength / 1000, 0.01, -3.71439881519581E-02)
        SendKeys "~"
        swModel.ClearSelection2 True
        Set myDimension = swModel.parameter("D6@Sketch12")
        myDimension.SystemValue = egright / 1000
        End If
    End If
    
    swModel.ClearSelection2 True
    swModel.SketchManager.InsertSketch True
    boolstatus = swModel.Extension.SelectByID2("Sketch12", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
    swModel.ViewZoomtofit2

    ' cut extrude holes into rhs/shs
    If punch = 1 Then
        Set myFeature = swModel.FeatureManager.FeatureCut4(True, False, False, 2, 0, 0.01, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
        swModel.SelectionManager.EnableContourSelection = False
    Else
        Set myFeature = swModel.FeatureManager.FeatureCut4(True, False, False, 1, 0, 0.01, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
        swModel.SelectionManager.EnableContourSelection = False
    End If
    
    ' housekeeping and cleanup
    swModel.ViewDispRefplanes
    boolstatus = swModel.Extension.SelectByID2("Sketch11", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
    boolstatus = swModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, True, 0, Nothing, 0)
    swModel.BlankSketch
    swModel.ClearSelection2 True
    
    ' save function
    Dim folderPath As String
    
    ' Specify the path of the new folder
    folderPath = "C:\Users\gtarn\OneDrive\Documents\FC Racing\Laser Cut\" & sordernum
    
    ' Check if the folder doesn't already exist
    If Dir(folderPath, vbDirectory) = "" Then
        ' Create the folder
        MkDir folderPath
        MsgBox "The folder " & Chr(34) & sordernum & Chr(34) & " has been created successfully. Files will be created inside.", vbOKOnly, "FC Racing 2024 " & Chr(169)
    Else
        MsgBox "The folder " & Chr(34) & sordernum & Chr(34) & " already exists. Files will be created inside.", vbOKOnly, "FC Racing 2024 " & Chr(169)
    End If
    
    
    ' Determining save name
    If matInput Like "A" Then
        If punch = 1 Then
            If clengthnyc = 6 Then
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "AP" & sheight & swidth & ") " & "SP " & fclength & " CUT" & " R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "AP" & sheight & swidth & ") " & "SP " & fclength & " CUT"
                End If
            Else
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "AP" & sheight & swidth & ") " & "SP " & "R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "AP" & sheight & swidth & ") " & "SP"
                End If
            End If
        Else
            If clengthnyc = 6 Then
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "AP" & sheight & swidth & ") " & "DP " & fclength & " CUT" & " R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "AP" & sheight & swidth & ") " & "DP " & fclength & " CUT"
                End If
            Else
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "AP" & sheight & swidth & ") " & "DP " & "R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "AP" & sheight & swidth & ") " & "DP"
                End If
            End If
        End If
    ElseIf matInput Like "G" Then
        If punch = 1 Then
            If clengthnyc = 6 Then
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "GP" & sheight & swidth & ") " & "SP " & fclength & " CUT" & " R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "GP" & sheight & swidth & ") " & "SP " & fclength & " CUT"
                End If
            Else
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "GP" & sheight & swidth & ") " & "SP " & "R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "GP" & sheight & swidth & ") " & "SP"
                End If
            End If
        Else
            If clengthnyc = 6 Then
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "GP" & sheight & swidth & ") " & "DP " & fclength & " CUT" & " R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "GP" & sheight & swidth & ") " & "DP " & fclength & " CUT"
                End If
            Else
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "GP" & sheight & swidth & ") " & "DP " & "R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "GP" & sheight & swidth & ") " & "DP"
                End If
            End If
        End If
    ElseIf matInput Like "S" Then
        If punch = 1 Then
            If clengthnyc = 6 Then
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "SP" & sheight & swidth & ") " & "SP " & fclength & " CUT" & " R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "SP" & sheight & swidth & ") " & "SP " & fclength & " CUT"
                End If
            Else
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "SP" & sheight & swidth & ") " & "SP " & "R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "SP" & sheight & swidth & ") " & "SP"
                End If
            End If
        Else
            If clengthnyc = 6 Then
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "SP" & sheight & swidth & ") " & "DP " & fclength & " CUT" & " R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "SP" & sheight & swidth & ") " & "DP " & fclength & " CUT"
                End If
            Else
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "SP" & sheight & swidth & ") " & "DP " & "R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "SP" & sheight & swidth & ") " & "DP"
                End If
            End If
        End If
    Else ' matInput Like "P"
        If punch = 1 Then
            If clengthnyc = 6 Then
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "PP" & sheight & swidth & ") " & "SP " & fclength & " CUT" & " R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "PP" & sheight & swidth & ") " & "SP " & fclength & " CUT"
                End If
            Else
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "PP" & sheight & swidth & ") " & "SP " & "R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "PP" & sheight & swidth & ") " & "SP"
                End If
            End If
        Else
            If clengthnyc = 6 Then
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "PP" & sheight & swidth & ") " & "DP " & fclength & " CUT" & " R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "PP" & sheight & swidth & ") " & "DP " & fclength & " CUT"
                End If
            Else
                If rakenyc = 6 Then
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "PP" & sheight & swidth & ") " & "DP " & "R" & rakeInput
                Else
                    savename = folderPath & "\" & sordernum & " - " _
                    & cordername & " (" & "PP" & sheight & swidth & ") " & "DP"
                End If
            End If
        End If
    End If
    
    ' Save file
    Set swModel = swApp.ActiveDoc
    Dim partname As String
    Dim stepname As String
    Dim drawname As String

    
    partname = savename & ".SLDPRT"
    stepname = savename & " STEP WORKS" & ".STEP"
    drawname = savename & ".SLDDRW"
    
    swModel.SaveAs3 partname, 0, 0  'as document
    swModel.SaveAs3 stepname, 0, 0 'as step file
    
    If rakenyc = 7 Then
        ' create drawing from part doc
        Dim swSheetWidth As Double
        swSheetWidth = 0.3
        Dim swSheetHeight As Double
        swSheetHeight = 0.2
        Set swModel = swApp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2022\templates\Drawing.drwdot", 12, swSheetWidth, swSheetHeight)
        Dim swDrawing As DrawingDoc
        Set swDrawing = swModel
        Set swDrawing = swModel
        Dim swSheet As Sheet
        Set swSheet = swDrawing.GetCurrentSheet()
        swSheet.SetProperties2 12, 12, 1, 1, False, swSheetWidth, swSheetHeight, True
        swSheet.SetTemplateName "c:\programdata\solidworks\solidworks 2022\lang\english\sheetformat\fc racing laser - landscape template.slddrt"
        swSheet.ReloadTemplate True
        boolstatus = swModel.GenerateViewPaletteViews(partname)
        
        ' Get drawing views based on holeside input
        Dim myView As Object
        If hsInput = sheight Then
            swDrawing.DropDrawingViewFromPalette2 "*Top", 0.15, 0.12, 0
            swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            Set myView = swModel.CreateUnfoldedViewAt3(0.15, 0.0866, 0, False)
            swModel.ClearSelection2 True
            swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            boolstatus = swModel.ActivateView("Drawing View1")
            swModel.ClearSelection2 True
            boolstatus = swModel.ActivateView("Drawing View2")
            boolstatus = swModel.Extension.SelectByID2("Drawing View2", "DRAWINGVIEW", 0.15, 0.0822, 0, False, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Drawing View2", "DRAWINGVIEW", 0.15, 0.0822, 0, False, 0, Nothing, 0)
            Set myView = swModel.CreateUnfoldedViewAt3(0.15, 0.0517, 0, False)
            swModel.ClearSelection2 True
            
            Set myView = swModel.DropDrawingViewFromPalette2("*Top", 0.15, 0.166, 0)
            boolstatus = swModel.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            swModel.ClearSelection2 True
            boolstatus = swModel.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0.102391724015879, 0.165762244240678, 0, False, 0, Nothing, 0)
            boolstatus = swModel.ActivateView("Drawing View4")
            swModel.ClearSelection2 True
            
        ElseIf hsInput = swidth Then
            swDrawing.DropDrawingViewFromPalette2 "*Front", 0.15, 0.12, 0
            swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            Set myView = swModel.CreateUnfoldedViewAt3(0.15, 0.0866, 0, False)
            swModel.ClearSelection2 True
            swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            boolstatus = swModel.ActivateView("Drawing View1")
            swModel.ClearSelection2 True
            boolstatus = swModel.ActivateView("Drawing View2")
            boolstatus = swModel.Extension.SelectByID2("Drawing View2", "DRAWINGVIEW", 0.15, 0.0822, 0, False, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Drawing View2", "DRAWINGVIEW", 0.15, 0.0822, 0, False, 0, Nothing, 0)
            Set myView = swModel.CreateUnfoldedViewAt3(0.15, 0.0517, 0, False)
            
            Set myView = swModel.DropDrawingViewFromPalette2("*Front", 0.15, 0.166, 0)
            boolstatus = swModel.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            swModel.ClearSelection2 True
            boolstatus = swModel.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0.102391724015879, 0.165762244240678, 0, False, 0, Nothing, 0)
            boolstatus = swModel.ActivateView("Drawing View4")
    
            'Dim myBreakLine As Object
            'Set myView = swModel.SelectionManager.GetSelectedObject5(1)
            'myBreakLine = myView.InsertBreak(0, -5.39328883740667E-02, 5.75514374219433E-02, 2)
            'swModel.BreakView
            
        Else 'hsInput = 0
            swDrawing.DropDrawingViewFromPalette2 "*Top", 0.15, 0.12, 0
            swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            Set myView = swModel.CreateUnfoldedViewAt3(0.15, 0.0866, 0, False)
            swModel.ClearSelection2 True
            swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            boolstatus = swModel.ActivateView("Drawing View1")
            swModel.ClearSelection2 True
            boolstatus = swModel.ActivateView("Drawing View2")
            boolstatus = swModel.Extension.SelectByID2("Drawing View2", "DRAWINGVIEW", 0.15, 0.0822, 0, False, 0, Nothing, 0)
            boolstatus = swModel.Extension.SelectByID2("Drawing View2", "DRAWINGVIEW", 0.15, 0.0822, 0, False, 0, Nothing, 0)
            Set myView = swModel.CreateUnfoldedViewAt3(0.15, 0.0517, 0, False)
            
            Set myView = swModel.DropDrawingViewFromPalette2("*Top", 0.15, 0.166, 0)
            boolstatus = swModel.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            swModel.ClearSelection2 True
            boolstatus = swModel.Extension.SelectByID2("Drawing View4", "DRAWINGVIEW", 0.102391724015879, 0.165762244240678, 0, False, 0, Nothing, 0)
            boolstatus = swModel.ActivateView("Drawing View4")
            swModel.ClearSelection2 True
            
        End If
        
        
        ' set scale of bottom drawings
        Dim swDraw As SldWorks.DrawingDoc
        Dim swSelMgr As SldWorks.SelectionMgr
        Dim swView As SldWorks.View
        Dim vScaleRatio As Variant
        Dim bRet As Boolean
        Set swApp = Application.SldWorks
        Set swModel = swApp.ActiveDoc
        Set swDraw = swModel
        
        swModel.ActivateView ("Drawing View1")
        swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        
        vScaleRatio = swView.ScaleRatio
        swView.ScaleDecimal = swView.ScaleDecimal * 2
        vScaleRatio = swView.ScaleRatio
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        ' set shaded for bottom drawings
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        swModel.ClearSelection2 True
        swModel.ActivateView ("Drawing View3")
        swModel.Extension.SelectByID2 "Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        swView.SetDisplayMode4 False, swSHADED, False, True, False
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        swModel.ClearSelection2 True
        swModel.ActivateView ("Drawing View2")
        swModel.Extension.SelectByID2 "Drawing View2", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        swView.SetDisplayMode4 False, swSHADED, False, True, False
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        swModel.ClearSelection2 True
        swModel.ActivateView ("Drawing View1")
        swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        swView.SetDisplayMode4 False, swSHADED, False, True, False
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        swModel.ClearSelection2 True
        swModel.ActivateView ("Drawing View4")
        swModel.Extension.SelectByID2 "Drawing View4", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        swView.SetDisplayMode4 False, swSHADED, False, True, False
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        ' adding text to drawing
        swDraw.CreateText2 quantity, 0.207, 0.025, 0, 0.0035, 0 'quantity
        
        If matInput Like "A" Then 'material
            swDraw.CreateText2 "ALU", 0.2045, 0.017, 0, 0.0035, 0
        ElseIf matInput Like "G" Then
            swDraw.CreateText2 "GAL", 0.2045, 0.017, 0, 0.0035, 0
        ElseIf matInput Like "S" Then
            swDraw.CreateText2 "SS", 0.2045, 0.017, 0, 0.0035, 0
        Else ' matInput Like "P" Then
            swDraw.CreateText2 "PAINTED", 0.195, 0.017, 0, 0.0035, 0
        End If
        
        swDraw.CreateText2 sordernum, 0.25, 0.017, 0, 0.0035, 0 'order number
        
        If clengthnyc = 6 Then 'cut length
            swDraw.CreateText2 fclength & " Cut", 0.25, 0.025, 0, 0.0035, 0
        Else
        End If
        
        swModel.EditRebuild3
        swModel.ForceRebuild3 False
        
        
        'saving drawing
        swModel.SaveAs3 drawname, 0, 0
    
    Else 'rakenyc = 6
        ' create drawing from part doc
        swSheetWidth = 0.3
        swSheetHeight = 0.2
        Set swModel = swApp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2022\templates\Drawing.drwdot", 12, swSheetWidth, swSheetHeight)
        Set swDrawing = swModel
        Set swSheet = swDrawing.GetCurrentSheet()
        swSheet.SetProperties2 12, 12, 1, 1, False, swSheetWidth, swSheetHeight, True
        swSheet.SetTemplateName "c:\programdata\solidworks\solidworks 2022\lang\english\sheetformat\FC Racing laser - landscape template raked.slddrt"
        swSheet.ReloadTemplate True
        boolstatus = swModel.GenerateViewPaletteViews(partname)
        
        If hsInput = sheight Or hsInput = 0 Then
            swDrawing.DropDrawingViewFromPalette2 "*Top", 0.15, 0.11, 0
            swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            swModel.ClearSelection2 True
            
            Set myView = swModel.DropDrawingViewFromPalette2("*Dimetric", 0.15, 0.05, 0)
            swDrawing.DropDrawingViewFromPalette2 "*Front", 0.15, 0.17, 0

            
        Else 'If hsInput = swidth Then
            swDrawing.DropDrawingViewFromPalette2 "*Front", 0.15, 0.11, 0
            swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            swModel.ClearSelection2 True
            
            Set myView = swModel.DropDrawingViewFromPalette2("*Dimetric", 0.15, 0.05, 0)
            swDrawing.DropDrawingViewFromPalette2 "*Top", 0.15, 0.17, 0
            
        End If

        Set swApp = Application.SldWorks
        Set swModel = swApp.ActiveDoc
        Set swDraw = swModel
        
        ' set scale of drawings
        swModel.ActivateView ("Drawing View1")
        swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        
        vScaleRatio = swView.ScaleRatio
        swView.ScaleDecimal = swView.ScaleDecimal * 2
        vScaleRatio = swView.ScaleRatio
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        swModel.ActivateView ("Drawing View2")
        swModel.Extension.SelectByID2 "Drawing View2", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        
        vScaleRatio = swView.ScaleRatio
        swView.ScaleDecimal = swView.ScaleDecimal * 2
        vScaleRatio = swView.ScaleRatio
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        swModel.ActivateView ("Drawing View3")
        swModel.Extension.SelectByID2 "Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        
        vScaleRatio = swView.ScaleRatio
        swView.ScaleDecimal = swView.ScaleDecimal * 2
        vScaleRatio = swView.ScaleRatio
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        
        ' set shaded for bottom drawings
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        swModel.ClearSelection2 True
        swModel.ActivateView ("Drawing View1")
        swModel.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        swView.SetDisplayMode4 False, swSHADED, False, True, False
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        swModel.ClearSelection2 True
        swModel.ActivateView ("Drawing View2")
        swModel.Extension.SelectByID2 "Drawing View2", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        swView.SetDisplayMode4 False, swSHADED, False, True, False
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        swModel.ClearSelection2 True
        swModel.ActivateView ("Drawing View3")
        swModel.Extension.SelectByID2 "Drawing View3", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set swSelMgr = swModel.SelectionManager
        Set swView = swSelMgr.GetSelectedObject6(1, -1)
        swView.SetDisplayMode4 False, swSHADED, False, True, False
        bRet = swModel.EditRebuild3: Debug.Assert bRet
        
        
        ' adding text to drawing
        swDraw.CreateText2 quantity, 0.207, 0.025, 0, 0.0035, 0 'quantity
        
        If matInput Like "A" Then 'material
            swDraw.CreateText2 "ALU", 0.2045, 0.017, 0, 0.0035, 0
        ElseIf matInput Like "G" Then
            swDraw.CreateText2 "GAL", 0.2045, 0.017, 0, 0.0035, 0
        ElseIf matInput Like "S" Then
            swDraw.CreateText2 "SS", 0.2045, 0.017, 0, 0.0035, 0
        Else ' matInput Like "P" Then
            swDraw.CreateText2 "PAINTED", 0.195, 0.017, 0, 0.0035, 0
        End If
        
        swDraw.CreateText2 sordernum, 0.25, 0.017, 0, 0.0035, 0 'order number
        
        ' rake and cut length
        If clengthnyc = 6 Then
            swDraw.CreateText2 fclength & " Cut, " & "Rake " & rakeInput & Chr(186), 0.23, 0.025, 0, 0.0035, 0
        Else 'clengthnyc = 7
            swDraw.CreateText2 "Rake " & rakeInput & Chr(186), 0.25, 0.025, 0, 0.0035, 0
        End If
        
        swModel.EditRebuild3
        swModel.ForceRebuild3 False
        
        
        'saving drawing
        swModel.SaveAs3 drawname, 0, 0
        
    End If
    
    
End Sub
Function SetParameters(prompt As String, ByRef parameter As Single) As Boolean
    Do
        genInput = InputBox("Please enter " & prompt & ":", "FC Racing 2024 " & Chr(169))

        If StrPtr(genInput) = 0 Then
            MsgBox "Input cancelled. Exiting."
            SetParameters = False ' User pressed "Cancel"
            Exit Function
        End If

        parameter = Val(genInput)

        If parameter < 0 Then
            MsgBox "Please enter positive values only."
        ElseIf parameter = 0 Then
            MsgBox "Please enter a non-zero value for " & prompt & "."
        End If
    Loop Until parameter > 0
    SetParameters = True
End Function





