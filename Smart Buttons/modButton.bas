Attribute VB_Name = "modButton"
Option Explicit

Public Sub UpdateButton(Op As Byte, oB As Object) 'This sub gets the Operation that you want to
'perform and the object that you want to perform it on.

    Dim oC As Control
    
        If TypeOf oB Is Form Then 'if the object you passed is a form then:
        
            For Each oC In oB.Controls 'Cycle through every control on the form
                
                If TypeOf oC Is Label Then 'If the current Control is a label (I use labels as buttons) then:
                
                    If oC.Enabled = True Then  'If the control is enabled then :
                        
                        If oC.BackColor = &HD63094 Then 'If the current backcolor of the control is &HD63094 then:
                            
                            'This step isn't required, but this prevents the sub from changing every label
                            'on the form's color, and to change only those of a specific color that you use for your
                            'buttons.
                            
                            oC.BackColor = &HFEE2E0
                            oC.ForeColor = vbBlack
                            
                        End If
                    
                    End If
                    
                End If
                
            Next
            
        Else 'If the type of object is not a form then:
        
            If Op = 1 Then 'If Op 1 is specified then:
            
            'Here we set the properties directly for the specified control because changes will
            'only be applied to this form and no other form. The values set is determined by
            'the value of Op that is passed to the sub.
            
                oB.BackColor = &HD63094
                oB.ForeColor = vbWhite
                
            Else
            
                oB.BackColor = &HFEE2E0
                oB.ForeColor = vbBlack
                
            End If
            
        End If

End Sub
