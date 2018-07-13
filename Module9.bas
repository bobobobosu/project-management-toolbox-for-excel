Attribute VB_Name = "Module9"
Sub waeiforunv()

            UserForm1.Show 'vbModeless
            FirstCell2 = Format(Range("ам╤у!I2").Value2, "m/d/yy")
            Dim FirstCell3 As Variant
            FirstCell3 = InputBox("Time Value", "Please Enter Time Value", "")
            
            MsgBox dateValue(FirstCell2) + TimeValue(FirstCell3)

End Sub
Filter
