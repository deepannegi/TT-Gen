Imports Microsoft.Office.Interop

Public Class frmIT_5A_G2
    Dim ColorFlag As Boolean = True        ' A Boolean variable for Color ON/OFF
    Private Sub CheckCorrectness() Handles CheckCorrectnessToolStripMenuItem.Click
        If cmb1.SelectedIndex < 0 Or cmb11.SelectedIndex < 0 Then
            cmb1.SelectedIndex = -1
            cmb11.SelectedIndex = -1
            pnl1.BackColor = pnlref.BackColor
            pnl11.BackColor = pnlref.BackColor
        End If
        If cmb2.SelectedIndex < 0 Or cmb12.SelectedIndex < 0 Then
            cmb2.SelectedIndex = -1
            cmb12.SelectedIndex = -1
            pnl2.BackColor = pnlref.BackColor
            pnl12.BackColor = pnlref.BackColor
        End If
        If cmb3.SelectedIndex < 0 Or cmb13.SelectedIndex < 0 Then
            cmb3.SelectedIndex = -1
            cmb13.SelectedIndex = -1
            pnl3.BackColor = pnlref.BackColor
            pnl13.BackColor = pnlref.BackColor
        End If
        If cmb4.SelectedIndex < 0 Or cmb14.SelectedIndex < 0 Then
            cmb4.SelectedIndex = -1
            cmb14.SelectedIndex = -1
            pnl4.BackColor = pnlref.BackColor
            pnl14.BackColor = pnlref.BackColor
        End If
        If cmb5.SelectedIndex < 0 Or cmb15.SelectedIndex < 0 Then
            cmb5.SelectedIndex = -1
            cmb15.SelectedIndex = -1
            pnl5.BackColor = pnlref.BackColor
            pnl15.BackColor = pnlref.BackColor
        End If
        If cmb6.SelectedIndex < 0 Or cmb16.SelectedIndex < 0 Then
            cmb6.SelectedIndex = -1
            cmb16.SelectedIndex = -1
            pnl6.BackColor = pnlref.BackColor
            pnl16.BackColor = pnlref.BackColor
        End If
        If cmb7.SelectedIndex < 0 Or cmb17.SelectedIndex < 0 Then
            cmb7.SelectedIndex = -1
            cmb17.SelectedIndex = -1
            pnl7.BackColor = pnlref.BackColor
            pnl17.BackColor = pnlref.BackColor
        End If
        If cmb8.SelectedIndex < 0 Or cmb18.SelectedIndex < 0 Then
            cmb8.SelectedIndex = -1
            cmb18.SelectedIndex = -1
            pnl8.BackColor = pnlref.BackColor
            pnl18.BackColor = pnlref.BackColor
        End If
        If cmb9.SelectedIndex < 0 Or cmb19.SelectedIndex < 0 Then
            cmb9.SelectedIndex = -1
            cmb19.SelectedIndex = -1
            pnl9.BackColor = pnlref.BackColor
            pnl19.BackColor = pnlref.BackColor
        End If
        If cmb10.SelectedIndex < 0 Or cmb20.SelectedIndex < 0 Then
            cmb10.SelectedIndex = -1
            cmb20.SelectedIndex = -1
            pnl10.BackColor = pnlref.BackColor
            pnl20.BackColor = pnlref.BackColor
        End If
        If cmb21.SelectedIndex < 0 Or cmb31.SelectedIndex < 0 Then
            cmb21.SelectedIndex = -1
            cmb31.SelectedIndex = -1
            pnl21.BackColor = pnlref.BackColor
            pnl31.BackColor = pnlref.BackColor
        End If
        If cmb22.SelectedIndex < 0 Or cmb32.SelectedIndex < 0 Then
            cmb22.SelectedIndex = -1
            cmb32.SelectedIndex = -1
            pnl22.BackColor = pnlref.BackColor
            pnl32.BackColor = pnlref.BackColor
        End If
        If cmb23.SelectedIndex < 0 Or cmb33.SelectedIndex < 0 Then
            cmb23.SelectedIndex = -1
            cmb33.SelectedIndex = -1
            pnl23.BackColor = pnlref.BackColor
            pnl33.BackColor = pnlref.BackColor
        End If
        If cmb24.SelectedIndex < 0 Or cmb34.SelectedIndex < 0 Then
            cmb24.SelectedIndex = -1
            cmb34.SelectedIndex = -1
            pnl24.BackColor = pnlref.BackColor
            pnl34.BackColor = pnlref.BackColor
        End If
        If cmb25.SelectedIndex < 0 Or cmb35.SelectedIndex < 0 Then
            cmb25.SelectedIndex = -1
            cmb35.SelectedIndex = -1
            pnl25.BackColor = pnlref.BackColor
            pnl35.BackColor = pnlref.BackColor
        End If
        If cmb26.SelectedIndex < 0 Or cmb36.SelectedIndex < 0 Then
            cmb26.SelectedIndex = -1
            cmb36.SelectedIndex = -1
            pnl26.BackColor = pnlref.BackColor
            pnl36.BackColor = pnlref.BackColor
        End If
        If cmb27.SelectedIndex < 0 Or cmb37.SelectedIndex < 0 Then
            cmb27.SelectedIndex = -1
            cmb37.SelectedIndex = -1
            pnl27.BackColor = pnlref.BackColor
            pnl37.BackColor = pnlref.BackColor
        End If
        If cmb28.SelectedIndex < 0 Or cmb38.SelectedIndex < 0 Then
            cmb28.SelectedIndex = -1
            cmb38.SelectedIndex = -1
            pnl28.BackColor = pnlref.BackColor
            pnl38.BackColor = pnlref.BackColor
        End If
        If cmb29.SelectedIndex < 0 Or cmb39.SelectedIndex < 0 Then
            cmb29.SelectedIndex = -1
            cmb39.SelectedIndex = -1
            pnl29.BackColor = pnlref.BackColor
            pnl39.BackColor = pnlref.BackColor
        End If
        If cmb30.SelectedIndex < 0 Or cmb40.SelectedIndex < 0 Then
            cmb30.SelectedIndex = -1
            cmb40.SelectedIndex = -1
            pnl30.BackColor = pnlref.BackColor
            pnl40.BackColor = pnlref.BackColor
        End If
        If cmb41.SelectedIndex < 0 Or cmb51.SelectedIndex < 0 Then
            cmb41.SelectedIndex = -1
            cmb51.SelectedIndex = -1
            pnl41.BackColor = pnlref.BackColor
            pnl51.BackColor = pnlref.BackColor
        End If
        If cmb42.SelectedIndex < 0 Or cmb52.SelectedIndex < 0 Then
            cmb42.SelectedIndex = -1
            cmb52.SelectedIndex = -1
            pnl42.BackColor = pnlref.BackColor
            pnl52.BackColor = pnlref.BackColor
        End If
        If cmb43.SelectedIndex < 0 Or cmb53.SelectedIndex < 0 Then
            cmb43.SelectedIndex = -1
            cmb53.SelectedIndex = -1
            pnl43.BackColor = pnlref.BackColor
            pnl53.BackColor = pnlref.BackColor
        End If
        If cmb44.SelectedIndex < 0 Or cmb54.SelectedIndex < 0 Then
            cmb44.SelectedIndex = -1
            cmb54.SelectedIndex = -1
            pnl44.BackColor = pnlref.BackColor
            pnl54.BackColor = pnlref.BackColor
        End If
        If cmb45.SelectedIndex < 0 Or cmb55.SelectedIndex < 0 Then
            cmb45.SelectedIndex = -1
            cmb55.SelectedIndex = -1
            pnl45.BackColor = pnlref.BackColor
            pnl55.BackColor = pnlref.BackColor
        End If
        If cmb46.SelectedIndex < 0 Or cmb56.SelectedIndex < 0 Then
            cmb46.SelectedIndex = -1
            cmb56.SelectedIndex = -1
            pnl46.BackColor = pnlref.BackColor
            pnl56.BackColor = pnlref.BackColor
        End If
        If cmb47.SelectedIndex < 0 Or cmb57.SelectedIndex < 0 Then
            cmb47.SelectedIndex = -1
            cmb57.SelectedIndex = -1
            pnl47.BackColor = pnlref.BackColor
            pnl57.BackColor = pnlref.BackColor
        End If
        If cmb48.SelectedIndex < 0 Or cmb58.SelectedIndex < 0 Then
            cmb48.SelectedIndex = -1
            cmb58.SelectedIndex = -1
            pnl48.BackColor = pnlref.BackColor
            pnl58.BackColor = pnlref.BackColor
        End If
        If cmb49.SelectedIndex < 0 Or cmb59.SelectedIndex < 0 Then
            cmb49.SelectedIndex = -1
            cmb59.SelectedIndex = -1
            pnl49.BackColor = pnlref.BackColor
            pnl59.BackColor = pnlref.BackColor
        End If
        If cmb50.SelectedIndex < 0 Or cmb60.SelectedIndex < 0 Then
            cmb50.SelectedIndex = -1
            cmb60.SelectedIndex = -1
            pnl50.BackColor = pnlref.BackColor
            pnl60.BackColor = pnlref.BackColor
        End If
        If cmb61.SelectedIndex < 0 Or cmb71.SelectedIndex < 0 Then
            cmb61.SelectedIndex = -1
            cmb71.SelectedIndex = -1
            pnl61.BackColor = pnlref.BackColor
            pnl71.BackColor = pnlref.BackColor
        End If
        If cmb62.SelectedIndex < 0 Or cmb72.SelectedIndex < 0 Then
            cmb62.SelectedIndex = -1
            cmb72.SelectedIndex = -1
            pnl62.BackColor = pnlref.BackColor
            pnl72.BackColor = pnlref.BackColor
        End If
        If cmb63.SelectedIndex < 0 Or cmb73.SelectedIndex < 0 Then
            cmb63.SelectedIndex = -1
            cmb73.SelectedIndex = -1
            pnl63.BackColor = pnlref.BackColor
            pnl73.BackColor = pnlref.BackColor
        End If
        If cmb64.SelectedIndex < 0 Or cmb74.SelectedIndex < 0 Then
            cmb64.SelectedIndex = -1
            cmb74.SelectedIndex = -1
            pnl64.BackColor = pnlref.BackColor
            pnl74.BackColor = pnlref.BackColor
        End If
        If cmb65.SelectedIndex < 0 Or cmb75.SelectedIndex < 0 Then
            cmb65.SelectedIndex = -1
            cmb75.SelectedIndex = -1
            pnl65.BackColor = pnlref.BackColor
            pnl75.BackColor = pnlref.BackColor
        End If
        If cmb66.SelectedIndex < 0 Or cmb76.SelectedIndex < 0 Then
            cmb66.SelectedIndex = -1
            cmb76.SelectedIndex = -1
            pnl66.BackColor = pnlref.BackColor
            pnl76.BackColor = pnlref.BackColor
        End If
        If cmb67.SelectedIndex < 0 Or cmb77.SelectedIndex < 0 Then
            cmb67.SelectedIndex = -1
            cmb77.SelectedIndex = -1
            pnl67.BackColor = pnlref.BackColor
            pnl77.BackColor = pnlref.BackColor
        End If
        If cmb68.SelectedIndex < 0 Or cmb78.SelectedIndex < 0 Then
            cmb68.SelectedIndex = -1
            cmb78.SelectedIndex = -1
            pnl68.BackColor = pnlref.BackColor
            pnl78.BackColor = pnlref.BackColor
        End If
        If cmb69.SelectedIndex < 0 Or cmb79.SelectedIndex < 0 Then
            cmb69.SelectedIndex = -1
            cmb79.SelectedIndex = -1
            pnl69.BackColor = pnlref.BackColor
            pnl79.BackColor = pnlref.BackColor
        End If
        If cmb70.SelectedIndex < 0 Or cmb80.SelectedIndex < 0 Then
            cmb70.SelectedIndex = -1
            cmb80.SelectedIndex = -1
            pnl70.BackColor = pnlref.BackColor
            pnl80.BackColor = pnlref.BackColor
        End If
        If cmb81.SelectedIndex < 0 Or cmb91.SelectedIndex < 0 Then
            cmb81.SelectedIndex = -1
            cmb91.SelectedIndex = -1
            pnl81.BackColor = pnlref.BackColor
            pnl91.BackColor = pnlref.BackColor
        End If
        If cmb82.SelectedIndex < 0 Or cmb92.SelectedIndex < 0 Then
            cmb82.SelectedIndex = -1
            cmb92.SelectedIndex = -1
            pnl82.BackColor = pnlref.BackColor
            pnl92.BackColor = pnlref.BackColor
        End If
        If cmb83.SelectedIndex < 0 Or cmb93.SelectedIndex < 0 Then
            cmb83.SelectedIndex = -1
            cmb93.SelectedIndex = -1
            pnl83.BackColor = pnlref.BackColor
            pnl93.BackColor = pnlref.BackColor
        End If
        If cmb84.SelectedIndex < 0 Or cmb94.SelectedIndex < 0 Then
            cmb84.SelectedIndex = -1
            cmb94.SelectedIndex = -1
            pnl84.BackColor = pnlref.BackColor
            pnl94.BackColor = pnlref.BackColor
        End If
        If cmb85.SelectedIndex < 0 Or cmb95.SelectedIndex < 0 Then
            cmb85.SelectedIndex = -1
            cmb95.SelectedIndex = -1
            pnl85.BackColor = pnlref.BackColor
            pnl95.BackColor = pnlref.BackColor
        End If
        If cmb86.SelectedIndex < 0 Or cmb96.SelectedIndex < 0 Then
            cmb86.SelectedIndex = -1
            cmb96.SelectedIndex = -1
            pnl86.BackColor = pnlref.BackColor
            pnl96.BackColor = pnlref.BackColor
        End If
        If cmb87.SelectedIndex < 0 Or cmb97.SelectedIndex < 0 Then
            cmb87.SelectedIndex = -1
            cmb97.SelectedIndex = -1
            pnl87.BackColor = pnlref.BackColor
            pnl97.BackColor = pnlref.BackColor
        End If
        If cmb88.SelectedIndex < 0 Or cmb98.SelectedIndex < 0 Then
            cmb88.SelectedIndex = -1
            cmb98.SelectedIndex = -1
            pnl88.BackColor = pnlref.BackColor
            pnl98.BackColor = pnlref.BackColor
        End If
        If cmb89.SelectedIndex < 0 Or cmb99.SelectedIndex < 0 Then
            cmb89.SelectedIndex = -1
            cmb99.SelectedIndex = -1
            pnl89.BackColor = pnlref.BackColor
            pnl99.BackColor = pnlref.BackColor
        End If
        If cmb90.SelectedIndex < 0 Or cmb100.SelectedIndex < 0 Then
            cmb90.SelectedIndex = -1
            cmb100.SelectedIndex = -1
            pnl90.BackColor = pnlref.BackColor
            pnl100.BackColor = pnlref.BackColor
        End If
    End Sub
    Private Sub RefreshColor()
        If ColorFlag Then
            ' Green Yellow Cream Red
            Dim count = 0
            If cmb1.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl1.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl1.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl1.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl1.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl1.BackColor = pnlref.BackColor
                    cmb1.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl1.BackColor = pnlref.BackColor
            End If
            If cmb2.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl2.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl2.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl2.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl2.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl2.BackColor = pnlref.BackColor
                    cmb2.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl2.BackColor = pnlref.BackColor
            End If

            If cmb3.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl3.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl3.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl3.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl3.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl3.BackColor = pnlref.BackColor
                    cmb3.SelectedIndex = -1
                    cmb3.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl3.BackColor = pnlref.BackColor
            End If
            If cmb4.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl4.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl4.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl4.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl4.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl4.BackColor = pnlref.BackColor
                    cmb4.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl4.BackColor = pnlref.BackColor
            End If
            If cmb5.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl5.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl5.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl5.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl5.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl5.BackColor = pnlref.BackColor
                    cmb5.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl5.BackColor = pnlref.BackColor
            End If
            If cmb6.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl6.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl6.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl6.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl6.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl6.BackColor = pnlref.BackColor
                    cmb6.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl6.BackColor = pnlref.BackColor
            End If
            If cmb7.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl7.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl7.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl7.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl7.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl7.BackColor = pnlref.BackColor
                    cmb7.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl7.BackColor = pnlref.BackColor
            End If
            If cmb8.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl8.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl8.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl8.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl8.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    cmb8.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl8.BackColor = pnlref.BackColor
            End If
            If cmb9.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl9.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl9.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl9.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl9.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl9.BackColor = pnlref.BackColor
                    cmb9.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl9.BackColor = pnlref.BackColor
            End If
            If cmb10.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl10.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl10.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl10.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl10.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl10.BackColor = pnlref.BackColor
                    cmb10.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl10.BackColor = pnlref.BackColor
            End If
            count = 0
            If cmb21.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl21.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl21.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl21.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl21.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl21.BackColor = pnlref.BackColor
                    cmb21.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl21.BackColor = pnlref.BackColor
            End If
            If cmb22.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl22.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl22.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl22.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl22.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl22.BackColor = pnlref.BackColor
                    cmb22.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl22.BackColor = pnlref.BackColor
            End If
            If cmb23.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl23.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl23.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl23.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl23.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl23.BackColor = pnlref.BackColor
                    cmb23.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl23.BackColor = pnlref.BackColor
            End If
            If cmb24.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl24.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl24.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl24.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl24.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl24.BackColor = pnlref.BackColor
                    cmb24.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl24.BackColor = pnlref.BackColor
            End If
            If cmb25.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl25.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl25.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl25.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl25.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl25.BackColor = pnlref.BackColor
                    cmb25.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl25.BackColor = pnlref.BackColor
            End If
            If cmb26.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl26.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl26.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl26.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl26.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl26.BackColor = pnlref.BackColor
                    cmb26.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl26.BackColor = pnlref.BackColor
            End If
            If cmb27.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl27.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl27.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl27.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl27.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl27.BackColor = pnlref.BackColor
                    cmb27.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl27.BackColor = pnlref.BackColor
            End If
            If cmb28.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl28.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl28.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl28.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl28.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl28.BackColor = pnlref.BackColor
                    cmb28.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl28.BackColor = pnlref.BackColor
            End If
            If cmb29.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl29.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl29.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl29.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl29.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl29.BackColor = pnlref.BackColor
                    cmb29.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl29.BackColor = pnlref.BackColor
            End If
            If cmb30.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl30.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl30.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl30.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl30.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl30.BackColor = pnlref.BackColor
                    cmb30.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl30.BackColor = pnlref.BackColor
            End If
            count = 0
            If cmb41.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl41.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl41.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl41.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl41.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl41.BackColor = pnlref.BackColor
                    cmb41.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl41.BackColor = pnlref.BackColor
            End If
            If cmb42.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl42.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl42.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl42.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl42.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl42.BackColor = pnlref.BackColor
                    cmb42.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl42.BackColor = pnlref.BackColor
            End If
            If cmb43.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl43.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl43.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl43.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl43.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl43.BackColor = pnlref.BackColor
                    cmb43.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl43.BackColor = pnlref.BackColor
            End If
            If cmb44.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl44.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl44.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl44.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl44.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl44.BackColor = pnlref.BackColor
                    cmb44.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl44.BackColor = pnlref.BackColor
            End If
            If cmb45.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl45.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl45.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl45.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl45.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl45.BackColor = pnlref.BackColor
                    cmb45.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl45.BackColor = pnlref.BackColor
            End If
            If cmb46.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl46.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl46.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl46.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl46.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl46.BackColor = pnlref.BackColor
                    cmb46.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl46.BackColor = pnlref.BackColor
            End If
            If cmb47.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl47.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl47.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl47.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl47.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl47.BackColor = pnlref.BackColor
                    cmb47.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl47.BackColor = pnlref.BackColor
            End If
            If cmb48.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl48.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl48.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl48.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl48.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl48.BackColor = pnlref.BackColor
                    cmb48.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl48.BackColor = pnlref.BackColor
            End If
            If cmb49.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl49.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl49.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl49.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl49.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl49.BackColor = pnlref.BackColor
                    cmb49.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl49.BackColor = pnlref.BackColor
            End If
            If cmb50.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl50.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl50.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl50.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl50.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl50.BackColor = pnlref.BackColor
                    cmb50.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl50.BackColor = pnlref.BackColor
            End If
            count = 0
            If cmb61.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl61.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl61.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl61.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl61.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl61.BackColor = pnlref.BackColor
                    cmb61.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl61.BackColor = pnlref.BackColor
            End If
            If cmb62.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl62.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl62.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl62.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl62.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl62.BackColor = pnlref.BackColor
                    cmb62.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl62.BackColor = pnlref.BackColor
            End If
            If cmb63.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl63.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl63.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl63.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl63.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl63.BackColor = pnlref.BackColor
                    cmb63.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl63.BackColor = pnlref.BackColor
            End If
            If cmb64.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl64.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl64.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl64.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl64.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl64.BackColor = pnlref.BackColor
                    cmb64.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl64.BackColor = pnlref.BackColor
            End If
            If cmb65.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl65.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl65.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl65.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl65.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl65.BackColor = pnlref.BackColor
                    cmb65.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl65.BackColor = pnlref.BackColor
            End If
            If cmb66.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl66.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl66.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl66.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl66.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl66.BackColor = pnlref.BackColor
                    cmb66.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl66.BackColor = pnlref.BackColor
            End If
            If cmb67.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl67.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl67.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl67.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl67.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl67.BackColor = pnlref.BackColor
                    cmb67.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl67.BackColor = pnlref.BackColor
            End If
            If cmb68.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl68.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl68.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl68.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl68.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl68.BackColor = pnlref.BackColor
                    cmb68.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl68.BackColor = pnlref.BackColor
            End If
            If cmb69.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl69.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl69.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl69.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl69.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl69.BackColor = pnlref.BackColor
                    cmb69.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl69.BackColor = pnlref.BackColor
            End If
            If cmb70.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl70.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl70.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl70.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl70.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl70.BackColor = pnlref.BackColor
                    cmb70.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl70.BackColor = pnlref.BackColor
            End If
            count = 0
            If cmb81.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl81.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl81.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl81.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl81.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl81.BackColor = pnlref.BackColor
                    cmb81.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl81.BackColor = pnlref.BackColor
            End If
            If cmb82.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl82.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl82.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl82.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl82.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl82.BackColor = pnlref.BackColor
                    cmb82.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl82.BackColor = pnlref.BackColor
            End If
            If cmb83.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl83.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl83.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl83.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl83.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl83.BackColor = pnlref.BackColor
                    cmb83.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl83.BackColor = pnlref.BackColor
            End If
            If cmb84.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl84.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl84.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl84.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl84.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl84.BackColor = pnlref.BackColor
                    cmb84.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl84.BackColor = pnlref.BackColor
            End If
            If cmb85.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl85.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl85.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl85.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl85.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl85.BackColor = pnlref.BackColor
                    cmb85.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl85.BackColor = pnlref.BackColor
            End If
            If cmb86.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl86.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl86.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl86.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl86.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl86.BackColor = pnlref.BackColor
                    cmb86.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl86.BackColor = pnlref.BackColor
            End If
            If cmb87.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl87.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl87.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl87.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl87.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl87.BackColor = pnlref.BackColor
                    cmb87.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl87.BackColor = pnlref.BackColor
            End If
            If cmb88.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl88.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl88.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl88.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl88.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl88.BackColor = pnlref.BackColor
                    cmb88.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl88.BackColor = pnlref.BackColor
            End If
            If cmb89.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl89.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl89.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl89.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl89.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl89.BackColor = pnlref.BackColor
                    cmb89.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl89.BackColor = pnlref.BackColor
            End If
            If cmb90.SelectedIndex >= 0 Then
                If count < 0 Then
                    count = 0
                End If
                If count = 0 Then
                    pnl90.BackColor = Color.Green
                ElseIf count = 1 Then
                    pnl90.BackColor = Color.GreenYellow
                ElseIf count = 2 Then
                    pnl90.BackColor = Color.LightPink
                ElseIf count = 3 Then
                    pnl90.BackColor = Color.Red
                Else
                    MsgBox("Time Table is hectic for students...!!!", vbCritical, vbOKOnly)
                    pnl90.BackColor = pnlref.BackColor
                    cmb90.SelectedIndex = -1
                End If
                count = count + 1
            Else
                count = count - 2
                pnl90.BackColor = pnlref.BackColor
            End If
        Else
            pnl1.BackColor = pnlref.BackColor
            pnl2.BackColor = pnlref.BackColor
            pnl3.BackColor = pnlref.BackColor
            pnl4.BackColor = pnlref.BackColor
            pnl5.BackColor = pnlref.BackColor
            pnl6.BackColor = pnlref.BackColor
            pnl7.BackColor = pnlref.BackColor
            pnl8.BackColor = pnlref.BackColor
            pnl9.BackColor = pnlref.BackColor

            pnl10.BackColor = pnlref.BackColor
            pnl11.BackColor = pnlref.BackColor
            pnl12.BackColor = pnlref.BackColor
            pnl13.BackColor = pnlref.BackColor
            pnl14.BackColor = pnlref.BackColor
            pnl15.BackColor = pnlref.BackColor
            pnl16.BackColor = pnlref.BackColor
            pnl17.BackColor = pnlref.BackColor
            pnl18.BackColor = pnlref.BackColor
            pnl19.BackColor = pnlref.BackColor

            pnl20.BackColor = pnlref.BackColor
            pnl21.BackColor = pnlref.BackColor
            pnl22.BackColor = pnlref.BackColor
            pnl23.BackColor = pnlref.BackColor
            pnl24.BackColor = pnlref.BackColor
            pnl25.BackColor = pnlref.BackColor
            pnl26.BackColor = pnlref.BackColor
            pnl27.BackColor = pnlref.BackColor
            pnl28.BackColor = pnlref.BackColor
            pnl29.BackColor = pnlref.BackColor

            pnl30.BackColor = pnlref.BackColor
            pnl31.BackColor = pnlref.BackColor
            pnl32.BackColor = pnlref.BackColor
            pnl33.BackColor = pnlref.BackColor
            pnl34.BackColor = pnlref.BackColor
            pnl35.BackColor = pnlref.BackColor
            pnl36.BackColor = pnlref.BackColor
            pnl37.BackColor = pnlref.BackColor
            pnl38.BackColor = pnlref.BackColor
            pnl39.BackColor = pnlref.BackColor

            pnl40.BackColor = pnlref.BackColor
            pnl41.BackColor = pnlref.BackColor
            pnl42.BackColor = pnlref.BackColor
            pnl43.BackColor = pnlref.BackColor
            pnl44.BackColor = pnlref.BackColor
            pnl45.BackColor = pnlref.BackColor
            pnl46.BackColor = pnlref.BackColor
            pnl47.BackColor = pnlref.BackColor
            pnl48.BackColor = pnlref.BackColor
            pnl49.BackColor = pnlref.BackColor

            pnl50.BackColor = pnlref.BackColor
            pnl51.BackColor = pnlref.BackColor
            pnl52.BackColor = pnlref.BackColor
            pnl53.BackColor = pnlref.BackColor
            pnl54.BackColor = pnlref.BackColor
            pnl55.BackColor = pnlref.BackColor
            pnl56.BackColor = pnlref.BackColor
            pnl57.BackColor = pnlref.BackColor
            pnl58.BackColor = pnlref.BackColor
            pnl59.BackColor = pnlref.BackColor

            pnl60.BackColor = pnlref.BackColor
            pnl61.BackColor = pnlref.BackColor
            pnl62.BackColor = pnlref.BackColor
            pnl63.BackColor = pnlref.BackColor
            pnl64.BackColor = pnlref.BackColor
            pnl65.BackColor = pnlref.BackColor
            pnl66.BackColor = pnlref.BackColor
            pnl67.BackColor = pnlref.BackColor
            pnl68.BackColor = pnlref.BackColor
            pnl69.BackColor = pnlref.BackColor

            pnl70.BackColor = pnlref.BackColor
            pnl71.BackColor = pnlref.BackColor
            pnl72.BackColor = pnlref.BackColor
            pnl73.BackColor = pnlref.BackColor
            pnl74.BackColor = pnlref.BackColor
            pnl75.BackColor = pnlref.BackColor
            pnl76.BackColor = pnlref.BackColor
            pnl77.BackColor = pnlref.BackColor
            pnl78.BackColor = pnlref.BackColor
            pnl79.BackColor = pnlref.BackColor

            pnl80.BackColor = pnlref.BackColor
            pnl81.BackColor = pnlref.BackColor
            pnl82.BackColor = pnlref.BackColor
            pnl83.BackColor = pnlref.BackColor
            pnl84.BackColor = pnlref.BackColor
            pnl85.BackColor = pnlref.BackColor
            pnl86.BackColor = pnlref.BackColor
            pnl87.BackColor = pnlref.BackColor
            pnl88.BackColor = pnlref.BackColor
            pnl89.BackColor = pnlref.BackColor

            pnl90.BackColor = pnlref.BackColor
            pnl91.BackColor = pnlref.BackColor
            pnl92.BackColor = pnlref.BackColor
            pnl93.BackColor = pnlref.BackColor
            pnl94.BackColor = pnlref.BackColor
            pnl95.BackColor = pnlref.BackColor
            pnl96.BackColor = pnlref.BackColor
            pnl97.BackColor = pnlref.BackColor
            pnl98.BackColor = pnlref.BackColor
            pnl99.BackColor = pnlref.BackColor

            pnl100.BackColor = pnlref.BackColor
        End If
    End Sub
    Private Sub RefreshList()
        frmMain.initDB()
        If cmb1.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb1.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb1.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb1.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb1.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb1.SelectedIndex = -1
            End If
        End If
        If cmb2.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb2.SelectedIndex = -1
            End If
        End If
        If cmb3.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb3.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb3.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb3.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb3.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb3.SelectedIndex = -1
            End If
        End If
        If cmb4.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb4.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb4.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb4.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb4.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb4.SelectedIndex = -1
            End If
        End If
        If cmb5.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb5.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb5.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb5.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb5.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb5.SelectedIndex = -1
            End If
        End If
        If cmb6.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb6.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb6.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb6.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb6.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb6.SelectedIndex = -1
            End If
        End If
        If cmb7.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb7.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb7.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb7.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb7.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb7.SelectedIndex = -1
            End If
        End If
        If cmb8.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb8.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb8.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb8.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb8.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb8.SelectedIndex = -1
            End If
        End If
        If cmb9.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb9.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb9.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb9.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb9.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb9.SelectedIndex = -1
            End If
        End If
        If cmb10.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb10.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb10.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb10.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb10.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb10.SelectedIndex = -1
            End If
        End If

        If cmb11.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb11.SelectedIndex)._M_8_9 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb11.SelectedIndex)._M_8_9 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb11.SelectedIndex = -1
            End If
        End If
        If cmb12.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb12.SelectedIndex)._M_9_10 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb12.SelectedIndex)._M_9_10 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb12.SelectedIndex = -1
            End If
        End If
        If cmb13.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb13.SelectedIndex)._M_10_11 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb13.SelectedIndex)._M_10_11 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb13.SelectedIndex = -1
            End If
        End If
        If cmb14.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb14.SelectedIndex)._M_11_12 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb14.SelectedIndex)._M_11_12 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb14.SelectedIndex = -1
            End If
        End If
        If cmb15.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb15.SelectedIndex)._M_12_1 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb15.SelectedIndex)._M_12_1 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb15.SelectedIndex = -1
            End If
        End If
        If cmb16.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb16.SelectedIndex)._M_1_2 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb16.SelectedIndex)._M_1_2 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb16.SelectedIndex = -1
            End If
        End If
        If cmb17.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb17.SelectedIndex)._M_2_3 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb17.SelectedIndex)._M_2_3 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb17.SelectedIndex = -1
            End If
        End If
        If cmb18.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb18.SelectedIndex)._M_3_4 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb18.SelectedIndex)._M_3_4 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb18.SelectedIndex = -1
            End If
        End If
        If cmb19.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb19.SelectedIndex)._M_4_5 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb19.SelectedIndex)._M_4_5 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb19.SelectedIndex = -1
            End If
        End If
        If cmb20.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb20.SelectedIndex)._M_5_6 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb20.SelectedIndex)._M_5_6 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb20.SelectedIndex = -1
            End If
        End If





        If cmb21.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb21.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb21.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb21.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb21.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb21.SelectedIndex = -1
            End If
        End If
        If cmb22.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb22.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb22.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb22.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb22.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb22.SelectedIndex = -1
            End If
        End If
        If cmb23.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb23.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb23.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb23.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb23.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb23.SelectedIndex = -1
            End If
        End If
        If cmb24.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb24.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb24.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb24.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb24.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb24.SelectedIndex = -1
            End If
        End If
        If cmb25.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb25.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb25.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb25.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb25.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb25.SelectedIndex = -1
            End If
        End If
        If cmb26.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb26.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb26.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb26.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb26.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb26.SelectedIndex = -1
            End If
        End If
        If cmb27.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb27.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb27.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb27.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb27.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb27.SelectedIndex = -1
            End If
        End If
        If cmb28.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb28.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb28.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb28.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb28.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb28.SelectedIndex = -1
            End If
        End If
        If cmb29.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb29.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb29.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb29.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb29.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb29.SelectedIndex = -1
            End If
        End If
        If cmb30.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb30.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb30.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb30.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb30.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb30.SelectedIndex = -1
            End If
        End If



        If cmb31.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb31.SelectedIndex)._TU_8_9 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb31.SelectedIndex)._TU_8_9 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb31.SelectedIndex = -1
            End If
        End If
        If cmb32.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb32.SelectedIndex)._TU_9_10 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb32.SelectedIndex)._TU_9_10 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb32.SelectedIndex = -1
            End If
        End If
        If cmb33.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb33.SelectedIndex)._TU_10_11 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb33.SelectedIndex)._TU_10_11 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb33.SelectedIndex = -1
            End If
        End If
        If cmb34.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb34.SelectedIndex)._TU_11_12 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb34.SelectedIndex)._TU_11_12 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb34.SelectedIndex = -1
            End If
        End If
        If cmb35.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb35.SelectedIndex)._TU_12_1 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb35.SelectedIndex)._TU_12_1 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb35.SelectedIndex = -1
            End If
        End If
        If cmb36.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb36.SelectedIndex)._TU_1_2 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb36.SelectedIndex)._TU_1_2 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb36.SelectedIndex = -1
            End If
        End If
        If cmb37.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb37.SelectedIndex)._TU_2_3 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb37.SelectedIndex)._TU_2_3 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb37.SelectedIndex = -1
            End If
        End If
        If cmb38.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb38.SelectedIndex)._TU_3_4 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb39.SelectedIndex)._TU_3_4 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb38.SelectedIndex = -1
            End If
        End If
        If cmb39.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb39.SelectedIndex)._TU_4_5 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb39.SelectedIndex)._TU_4_5 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb39.SelectedIndex = -1
            End If
        End If
        If cmb40.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb40.SelectedIndex)._TU_5_6 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb40.SelectedIndex)._TU_5_6 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb40.SelectedIndex = -1
            End If
        End If

        If cmb41.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb41.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb41.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb41.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb41.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb41.SelectedIndex = -1
            End If
        End If
        If cmb42.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb42.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb42.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb42.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb42.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb42.SelectedIndex = -1
            End If
        End If
        If cmb43.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb43.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb43.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb43.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb43.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb43.SelectedIndex = -1
            End If
        End If
        If cmb44.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb44.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb44.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb44.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb44.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb44.SelectedIndex = -1
            End If
        End If
        If cmb45.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb45.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb45.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb45.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb45.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb45.SelectedIndex = -1
            End If
        End If
        If cmb46.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb46.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb46.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb46.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb46.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb46.SelectedIndex = -1
            End If
        End If
        If cmb47.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb47.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb47.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb47.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb47.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb47.SelectedIndex = -1
            End If
        End If
        If cmb48.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb48.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb48.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb48.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb48.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb48.SelectedIndex = -1
            End If
        End If
        If cmb49.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb49.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb49.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb49.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb49.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb49.SelectedIndex = -1
            End If
        End If
        If cmb50.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb50.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb50.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb50.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb50.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb50.SelectedIndex = -1
            End If
        End If


        If cmb51.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb51.SelectedIndex)._TU_8_9 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb51.SelectedIndex)._TU_8_9 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb51.SelectedIndex = -1
            End If
        End If
        If cmb52.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb52.SelectedIndex)._TU_9_10 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb52.SelectedIndex)._TU_9_10 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb52.SelectedIndex = -1
            End If
        End If
        If cmb53.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb53.SelectedIndex)._TU_10_11 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb53.SelectedIndex)._TU_10_11 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb53.SelectedIndex = -1
            End If
        End If
        If cmb54.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb54.SelectedIndex)._TU_11_12 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb54.SelectedIndex)._TU_11_12 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb54.SelectedIndex = -1
            End If
        End If
        If cmb55.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb55.SelectedIndex)._TU_12_1 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb55.SelectedIndex)._TU_12_1 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb55.SelectedIndex = -1
            End If
        End If
        If cmb56.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb56.SelectedIndex)._TU_1_2 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb56.SelectedIndex)._TU_1_2 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb56.SelectedIndex = -1
            End If
        End If
        If cmb57.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb57.SelectedIndex)._TU_2_3 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb57.SelectedIndex)._TU_2_3 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb57.SelectedIndex = -1
            End If
        End If
        If cmb58.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb58.SelectedIndex)._TU_3_4 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb59.SelectedIndex)._TU_3_4 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb58.SelectedIndex = -1
            End If
        End If
        If cmb59.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb59.SelectedIndex)._TU_4_5 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb59.SelectedIndex)._TU_4_5 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb59.SelectedIndex = -1
            End If
        End If
        If cmb60.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb60.SelectedIndex)._TU_5_6 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb60.SelectedIndex)._TU_5_6 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb60.SelectedIndex = -1
            End If
        End If




        If cmb61.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb61.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb61.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb61.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb61.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb61.SelectedIndex = -1
            End If
        End If
        If cmb62.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb62.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb62.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb62.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb62.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb62.SelectedIndex = -1
            End If
        End If
        If cmb63.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb63.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb63.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb63.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb63.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb63.SelectedIndex = -1
            End If
        End If
        If cmb64.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb64.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb64.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb64.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb64.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb64.SelectedIndex = -1
            End If
        End If
        If cmb65.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb65.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb65.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb65.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb65.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb65.SelectedIndex = -1
            End If
        End If
        If cmb66.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb66.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb66.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb66.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb66.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb66.SelectedIndex = -1
            End If
        End If
        If cmb67.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb67.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb67.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb67.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb67.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb67.SelectedIndex = -1
            End If
        End If
        If cmb68.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb68.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb68.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb68.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb68.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb68.SelectedIndex = -1
            End If
        End If
        If cmb69.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb69.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb69.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb69.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb69.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb69.SelectedIndex = -1
            End If
        End If
        If cmb70.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb70.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb70.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb70.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb70.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb70.SelectedIndex = -1
            End If
        End If



        If cmb71.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb71.SelectedIndex)._TH_8_9 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb71.SelectedIndex)._TH_8_9 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb71.SelectedIndex = -1
            End If
        End If
        If cmb72.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb72.SelectedIndex)._TH_9_10 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb72.SelectedIndex)._TH_9_10 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb72.SelectedIndex = -1
            End If
        End If
        If cmb73.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb73.SelectedIndex)._TH_10_11 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb73.SelectedIndex)._TH_10_11 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb73.SelectedIndex = -1
            End If
        End If
        If cmb74.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb74.SelectedIndex)._TH_11_12 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb74.SelectedIndex)._TH_11_12 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb74.SelectedIndex = -1
            End If
        End If
        If cmb75.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb75.SelectedIndex)._TH_12_1 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb75.SelectedIndex)._TH_12_1 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb75.SelectedIndex = -1
            End If
        End If
        If cmb76.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb76.SelectedIndex)._TH_1_2 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb76.SelectedIndex)._TH_1_2 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb76.SelectedIndex = -1
            End If
        End If
        If cmb77.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb77.SelectedIndex)._TH_2_3 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb77.SelectedIndex)._TH_2_3 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb77.SelectedIndex = -1
            End If
        End If
        If cmb78.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb78.SelectedIndex)._TH_3_4 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb78.SelectedIndex)._TH_3_4 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb78.SelectedIndex = -1
            End If
        End If
        If cmb79.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb79.SelectedIndex)._TH_4_5 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb79.SelectedIndex)._TH_4_5 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb79.SelectedIndex = -1
            End If
        End If
        If cmb80.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb80.SelectedIndex)._TH_5_6 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb80.SelectedIndex)._TH_5_6 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb80.SelectedIndex = -1
            End If
        End If




        If cmb81.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb81.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb81.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb81.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb81.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb81.SelectedIndex = -1
            End If
        End If
        If cmb82.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb82.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb82.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb82.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb82.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb82.SelectedIndex = -1
            End If
        End If
        If cmb83.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb83.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb83.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb83.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb83.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb83.SelectedIndex = -1
            End If
        End If
        If cmb84.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb84.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb84.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb84.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb84.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb84.SelectedIndex = -1
            End If
        End If
        If cmb85.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb85.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb85.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb85.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb85.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb85.SelectedIndex = -1
            End If
        End If
        If cmb86.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb86.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb86.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb86.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb86.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb86.SelectedIndex = -1
            End If
        End If
        If cmb87.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb87.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb87.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb87.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb87.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb87.SelectedIndex = -1
            End If
        End If
        If cmb88.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb88.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb88.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb88.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb88.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb88.SelectedIndex = -1
            End If
        End If
        If cmb89.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb89.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb89.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb89.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb89.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb89.SelectedIndex = -1
            End If
        End If
        If cmb90.Text <> "" Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb90.SelectedIndex).Taken < frmMain.DB_oddDataSet1.IT_5A_G2(cmb90.SelectedIndex).No_per_week Then
                frmMain.DB_oddDataSet1.IT_5A_G2(cmb90.SelectedIndex).Taken = frmMain.DB_oddDataSet1.IT_5A_G2(cmb90.SelectedIndex).Taken + 1
            Else
                MsgBox("Limit Reached for this subject ...!!!")
                cmb90.SelectedIndex = -1
            End If
        End If



        If cmb91.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb91.SelectedIndex)._F_8_9 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb91.SelectedIndex)._F_8_9 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb91.SelectedIndex = -1
            End If
        End If
        If cmb92.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb92.SelectedIndex)._F_9_10 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb92.SelectedIndex)._F_9_10 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb92.SelectedIndex = -1
            End If
        End If
        If cmb93.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb93.SelectedIndex)._F_10_11 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb93.SelectedIndex)._F_10_11 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb93.SelectedIndex = -1
            End If
        End If
        If cmb94.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb94.SelectedIndex)._F_11_12 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb94.SelectedIndex)._F_11_12 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb94.SelectedIndex = -1
            End If
        End If
        If cmb95.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb95.SelectedIndex)._F_12_1 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb95.SelectedIndex)._F_12_1 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb95.SelectedIndex = -1
            End If
        End If
        If cmb96.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb96.SelectedIndex)._F_1_2 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb96.SelectedIndex)._F_1_2 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb96.SelectedIndex = -1
            End If
        End If
        If cmb97.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb97.SelectedIndex)._F_2_3 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb97.SelectedIndex)._F_2_3 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb97.SelectedIndex = -1
            End If
        End If
        If cmb98.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb98.SelectedIndex)._F_3_4 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb99.SelectedIndex)._F_3_4 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb98.SelectedIndex = -1
            End If
        End If
        If cmb99.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb99.SelectedIndex)._F_4_5 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb99.SelectedIndex)._F_4_5 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb99.SelectedIndex = -1
            End If
        End If
        If cmb40.Text <> "" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb40.SelectedIndex)._F_5_6 = False Then
                frmMain.DB_room_ODDDataSet1.Rooms(cmb40.SelectedIndex)._F_5_6 = True
            Else
                MsgBox("Room is preoccupied ...!!!")
                cmb40.SelectedIndex = -1
            End If
        End If
        RefreshColor()
    End Sub
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If MsgBox("Lectures have been coppied to other groups!!!") Then
            frmIT_5A_G1.Show()
            frmIT_5A_G3.Show()
        End If
        If cmb1.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb1.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb1.SelectedIndex = Me.cmb1.SelectedIndex
                frmIT_5A_G3.cmb1.SelectedIndex = Me.cmb1.SelectedIndex
            End If
        End If
        If cmb2.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb2.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb2.SelectedIndex = Me.cmb2.SelectedIndex
                frmIT_5A_G3.cmb2.SelectedIndex = Me.cmb2.SelectedIndex
            End If
        End If
        If cmb3.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb3.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb3.SelectedIndex = Me.cmb3.SelectedIndex
                frmIT_5A_G3.cmb3.SelectedIndex = Me.cmb3.SelectedIndex
            End If
        End If
        If cmb4.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb4.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb4.SelectedIndex = Me.cmb4.SelectedIndex
                frmIT_5A_G3.cmb4.SelectedIndex = Me.cmb4.SelectedIndex
            End If
        End If
        If cmb5.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb5.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb5.SelectedIndex = Me.cmb5.SelectedIndex
                frmIT_5A_G3.cmb5.SelectedIndex = Me.cmb5.SelectedIndex
            End If
        End If
        If cmb6.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb6.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb6.SelectedIndex = Me.cmb6.SelectedIndex
                frmIT_5A_G3.cmb6.SelectedIndex = Me.cmb6.SelectedIndex
            End If
        End If
        If cmb7.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb7.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb7.SelectedIndex = Me.cmb7.SelectedIndex
                frmIT_5A_G3.cmb7.SelectedIndex = Me.cmb7.SelectedIndex
            End If
        End If
        If cmb8.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb8.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb8.SelectedIndex = Me.cmb8.SelectedIndex
                frmIT_5A_G3.cmb8.SelectedIndex = Me.cmb8.SelectedIndex
            End If
        End If
        If cmb9.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb9.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb9.SelectedIndex = Me.cmb9.SelectedIndex
                frmIT_5A_G3.cmb9.SelectedIndex = Me.cmb9.SelectedIndex
            End If
        End If
        If cmb10.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb10.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb10.SelectedIndex = Me.cmb10.SelectedIndex
                frmIT_5A_G3.cmb10.SelectedIndex = Me.cmb10.SelectedIndex
            End If
        End If
        If cmb11.SelectedIndex >= 0 And cmb1.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb1.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb11.SelectedIndex = Me.cmb11.SelectedIndex
                frmIT_5A_G3.cmb11.SelectedIndex = Me.cmb11.SelectedIndex
            End If
        End If
        If cmb12.SelectedIndex >= 0 And cmb2.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb2.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb12.SelectedIndex = Me.cmb12.SelectedIndex
                frmIT_5A_G3.cmb12.SelectedIndex = Me.cmb12.SelectedIndex
            End If
        End If
        If cmb13.SelectedIndex >= 0 And cmb3.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb3.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb13.SelectedIndex = Me.cmb13.SelectedIndex
                frmIT_5A_G3.cmb13.SelectedIndex = Me.cmb13.SelectedIndex
            End If
        End If
        If cmb13.SelectedIndex >= 0 And cmb3.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb3.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb13.SelectedIndex = Me.cmb13.SelectedIndex
                frmIT_5A_G3.cmb13.SelectedIndex = Me.cmb13.SelectedIndex
            End If
        End If
        If cmb14.SelectedIndex >= 0 And cmb4.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb4.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb14.SelectedIndex = Me.cmb14.SelectedIndex
                frmIT_5A_G3.cmb14.SelectedIndex = Me.cmb14.SelectedIndex
            End If
        End If
        If cmb15.SelectedIndex >= 0 And cmb5.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb5.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb15.SelectedIndex = Me.cmb15.SelectedIndex
                frmIT_5A_G3.cmb15.SelectedIndex = Me.cmb15.SelectedIndex
            End If
        End If
        If cmb16.SelectedIndex >= 0 And cmb6.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb6.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb16.SelectedIndex = Me.cmb16.SelectedIndex
                frmIT_5A_G3.cmb16.SelectedIndex = Me.cmb16.SelectedIndex
            End If
        End If
        If cmb17.SelectedIndex >= 0 And cmb7.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb7.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb17.SelectedIndex = Me.cmb17.SelectedIndex
                frmIT_5A_G3.cmb17.SelectedIndex = Me.cmb17.SelectedIndex
            End If
        End If
        If cmb18.SelectedIndex >= 0 And cmb8.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb8.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb18.SelectedIndex = Me.cmb18.SelectedIndex
                frmIT_5A_G3.cmb18.SelectedIndex = Me.cmb18.SelectedIndex
            End If
        End If
        If cmb19.SelectedIndex >= 0 And cmb9.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb9.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb19.SelectedIndex = Me.cmb19.SelectedIndex
                frmIT_5A_G3.cmb19.SelectedIndex = Me.cmb19.SelectedIndex
            End If
        End If
        If cmb20.SelectedIndex >= 0 And cmb10.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb10.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb20.SelectedIndex = Me.cmb20.SelectedIndex
                frmIT_5A_G3.cmb20.SelectedIndex = Me.cmb20.SelectedIndex
            End If
        End If
        If cmb21.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb21.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb21.SelectedIndex = Me.cmb21.SelectedIndex
                frmIT_5A_G3.cmb21.SelectedIndex = Me.cmb21.SelectedIndex
            End If
        End If
        If cmb22.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb22.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb22.SelectedIndex = Me.cmb22.SelectedIndex
                frmIT_5A_G3.cmb22.SelectedIndex = Me.cmb22.SelectedIndex
            End If
        End If
        If cmb23.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb23.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb23.SelectedIndex = Me.cmb23.SelectedIndex
                frmIT_5A_G3.cmb23.SelectedIndex = Me.cmb23.SelectedIndex
            End If
        End If
        If cmb24.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb24.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb24.SelectedIndex = Me.cmb24.SelectedIndex
                frmIT_5A_G3.cmb24.SelectedIndex = Me.cmb24.SelectedIndex
            End If
        End If
        If cmb25.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb25.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb25.SelectedIndex = Me.cmb25.SelectedIndex
                frmIT_5A_G3.cmb25.SelectedIndex = Me.cmb25.SelectedIndex
            End If
        End If
        If cmb26.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb26.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb26.SelectedIndex = Me.cmb26.SelectedIndex
                frmIT_5A_G3.cmb26.SelectedIndex = Me.cmb26.SelectedIndex
            End If
        End If
        If cmb27.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb27.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb27.SelectedIndex = Me.cmb27.SelectedIndex
                frmIT_5A_G3.cmb27.SelectedIndex = Me.cmb27.SelectedIndex
            End If
        End If
        If cmb28.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb28.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb28.SelectedIndex = Me.cmb28.SelectedIndex
                frmIT_5A_G3.cmb28.SelectedIndex = Me.cmb28.SelectedIndex
            End If
        End If
        If cmb29.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb29.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb29.SelectedIndex = Me.cmb29.SelectedIndex
                frmIT_5A_G3.cmb29.SelectedIndex = Me.cmb29.SelectedIndex
            End If
        End If
        If cmb30.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb30.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb30.SelectedIndex = Me.cmb30.SelectedIndex
                frmIT_5A_G3.cmb30.SelectedIndex = Me.cmb30.SelectedIndex
            End If
        End If
        If cmb31.SelectedIndex >= 0 And cmb21.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb21.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb31.SelectedIndex = Me.cmb31.SelectedIndex
                frmIT_5A_G3.cmb31.SelectedIndex = Me.cmb31.SelectedIndex
            End If
        End If
        If cmb32.SelectedIndex >= 0 And cmb22.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb22.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb32.SelectedIndex = Me.cmb32.SelectedIndex
                frmIT_5A_G3.cmb32.SelectedIndex = Me.cmb32.SelectedIndex
            End If
        End If
        If cmb33.SelectedIndex >= 0 And cmb23.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb23.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb33.SelectedIndex = Me.cmb33.SelectedIndex
                frmIT_5A_G3.cmb33.SelectedIndex = Me.cmb33.SelectedIndex
            End If
        End If
        If cmb34.SelectedIndex >= 0 And cmb24.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb24.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb34.SelectedIndex = Me.cmb34.SelectedIndex
                frmIT_5A_G3.cmb34.SelectedIndex = Me.cmb34.SelectedIndex
            End If
        End If
        If cmb35.SelectedIndex >= 0 And cmb25.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb25.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb35.SelectedIndex = Me.cmb35.SelectedIndex
                frmIT_5A_G3.cmb35.SelectedIndex = Me.cmb35.SelectedIndex
            End If
        End If
        If cmb36.SelectedIndex >= 0 And cmb26.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb26.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb36.SelectedIndex = Me.cmb36.SelectedIndex
                frmIT_5A_G3.cmb36.SelectedIndex = Me.cmb36.SelectedIndex
            End If
        End If
        If cmb37.SelectedIndex >= 0 And cmb27.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb27.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb37.SelectedIndex = Me.cmb37.SelectedIndex
                frmIT_5A_G3.cmb37.SelectedIndex = Me.cmb37.SelectedIndex
            End If
        End If
        If cmb38.SelectedIndex >= 0 And cmb28.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb28.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb38.SelectedIndex = Me.cmb38.SelectedIndex
                frmIT_5A_G3.cmb38.SelectedIndex = Me.cmb38.SelectedIndex
            End If
        End If
        If cmb39.SelectedIndex >= 0 And cmb29.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb29.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb39.SelectedIndex = Me.cmb39.SelectedIndex
                frmIT_5A_G3.cmb39.SelectedIndex = Me.cmb39.SelectedIndex
            End If
        End If
        If cmb40.SelectedIndex >= 0 And cmb30.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb30.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb40.SelectedIndex = Me.cmb40.SelectedIndex
                frmIT_5A_G3.cmb40.SelectedIndex = Me.cmb40.SelectedIndex
            End If
        End If
        If cmb41.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb41.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb41.SelectedIndex = Me.cmb41.SelectedIndex
                frmIT_5A_G3.cmb41.SelectedIndex = Me.cmb41.SelectedIndex
            End If
        End If
        If cmb42.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb42.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb42.SelectedIndex = Me.cmb42.SelectedIndex
                frmIT_5A_G3.cmb42.SelectedIndex = Me.cmb42.SelectedIndex
            End If
        End If
        If cmb43.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb43.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb43.SelectedIndex = Me.cmb43.SelectedIndex
                frmIT_5A_G3.cmb43.SelectedIndex = Me.cmb43.SelectedIndex
            End If
        End If
        If cmb44.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb44.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb44.SelectedIndex = Me.cmb44.SelectedIndex
                frmIT_5A_G3.cmb44.SelectedIndex = Me.cmb44.SelectedIndex
            End If
        End If
        If cmb45.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb45.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb45.SelectedIndex = Me.cmb45.SelectedIndex
                frmIT_5A_G3.cmb45.SelectedIndex = Me.cmb45.SelectedIndex
            End If
        End If
        If cmb46.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb46.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb46.SelectedIndex = Me.cmb46.SelectedIndex
                frmIT_5A_G3.cmb46.SelectedIndex = Me.cmb46.SelectedIndex
            End If
        End If
        If cmb47.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb47.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb47.SelectedIndex = Me.cmb47.SelectedIndex
                frmIT_5A_G3.cmb47.SelectedIndex = Me.cmb47.SelectedIndex
            End If
        End If
        If cmb48.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb48.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb48.SelectedIndex = Me.cmb48.SelectedIndex
                frmIT_5A_G3.cmb48.SelectedIndex = Me.cmb48.SelectedIndex
            End If
        End If
        If cmb49.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb49.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb49.SelectedIndex = Me.cmb49.SelectedIndex
                frmIT_5A_G3.cmb49.SelectedIndex = Me.cmb49.SelectedIndex
            End If
        End If
        If cmb50.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb50.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb50.SelectedIndex = Me.cmb50.SelectedIndex
                frmIT_5A_G3.cmb50.SelectedIndex = Me.cmb50.SelectedIndex
            End If
        End If
        If cmb51.SelectedIndex >= 0 And cmb41.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb41.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb51.SelectedIndex = Me.cmb51.SelectedIndex
                frmIT_5A_G3.cmb51.SelectedIndex = Me.cmb51.SelectedIndex
            End If
        End If
        If cmb52.SelectedIndex >= 0 And cmb42.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb42.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb52.SelectedIndex = Me.cmb52.SelectedIndex
                frmIT_5A_G3.cmb52.SelectedIndex = Me.cmb52.SelectedIndex
            End If
        End If
        If cmb53.SelectedIndex >= 0 And cmb43.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb43.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb53.SelectedIndex = Me.cmb53.SelectedIndex
                frmIT_5A_G3.cmb53.SelectedIndex = Me.cmb53.SelectedIndex
            End If
        End If
        If cmb54.SelectedIndex >= 0 And cmb44.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb44.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb54.SelectedIndex = Me.cmb54.SelectedIndex
                frmIT_5A_G3.cmb54.SelectedIndex = Me.cmb54.SelectedIndex
            End If
        End If
        If cmb55.SelectedIndex >= 0 And cmb45.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb45.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb55.SelectedIndex = Me.cmb55.SelectedIndex
                frmIT_5A_G3.cmb55.SelectedIndex = Me.cmb55.SelectedIndex
            End If
        End If
        If cmb56.SelectedIndex >= 0 And cmb46.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb46.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb56.SelectedIndex = Me.cmb56.SelectedIndex
                frmIT_5A_G3.cmb56.SelectedIndex = Me.cmb56.SelectedIndex
            End If
        End If
        If cmb57.SelectedIndex >= 0 And cmb47.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb47.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb57.SelectedIndex = Me.cmb57.SelectedIndex
                frmIT_5A_G3.cmb57.SelectedIndex = Me.cmb57.SelectedIndex
            End If
        End If
        If cmb58.SelectedIndex >= 0 And cmb48.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb48.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb58.SelectedIndex = Me.cmb58.SelectedIndex
                frmIT_5A_G3.cmb58.SelectedIndex = Me.cmb58.SelectedIndex
            End If
        End If
        If cmb59.SelectedIndex >= 0 And cmb49.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb49.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb59.SelectedIndex = Me.cmb59.SelectedIndex
                frmIT_5A_G3.cmb59.SelectedIndex = Me.cmb59.SelectedIndex
            End If
        End If
        If cmb60.SelectedIndex >= 0 And cmb50.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb50.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb60.SelectedIndex = Me.cmb60.SelectedIndex
                frmIT_5A_G3.cmb60.SelectedIndex = Me.cmb60.SelectedIndex
            End If
        End If
        If cmb61.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb61.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb61.SelectedIndex = Me.cmb61.SelectedIndex
                frmIT_5A_G3.cmb61.SelectedIndex = Me.cmb61.SelectedIndex
            End If
        End If
        If cmb62.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb62.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb62.SelectedIndex = Me.cmb62.SelectedIndex
                frmIT_5A_G3.cmb62.SelectedIndex = Me.cmb62.SelectedIndex
            End If
        End If
        If cmb63.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb63.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb63.SelectedIndex = Me.cmb63.SelectedIndex
                frmIT_5A_G3.cmb63.SelectedIndex = Me.cmb63.SelectedIndex
            End If
        End If
        If cmb64.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb64.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb64.SelectedIndex = Me.cmb64.SelectedIndex
                frmIT_5A_G3.cmb64.SelectedIndex = Me.cmb64.SelectedIndex
            End If
        End If
        If cmb65.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb65.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb65.SelectedIndex = Me.cmb65.SelectedIndex
                frmIT_5A_G3.cmb65.SelectedIndex = Me.cmb65.SelectedIndex
            End If
        End If
        If cmb66.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb66.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb66.SelectedIndex = Me.cmb66.SelectedIndex
                frmIT_5A_G3.cmb66.SelectedIndex = Me.cmb66.SelectedIndex
            End If
        End If
        If cmb67.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb67.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb67.SelectedIndex = Me.cmb67.SelectedIndex
                frmIT_5A_G3.cmb67.SelectedIndex = Me.cmb67.SelectedIndex
            End If
        End If
        If cmb68.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb68.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb68.SelectedIndex = Me.cmb68.SelectedIndex
                frmIT_5A_G3.cmb68.SelectedIndex = Me.cmb68.SelectedIndex
            End If
        End If
        If cmb69.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb69.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb69.SelectedIndex = Me.cmb69.SelectedIndex
                frmIT_5A_G3.cmb69.SelectedIndex = Me.cmb69.SelectedIndex
            End If
        End If
        If cmb70.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb70.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb70.SelectedIndex = Me.cmb70.SelectedIndex
                frmIT_5A_G3.cmb70.SelectedIndex = Me.cmb70.SelectedIndex
            End If
        End If
        If cmb71.SelectedIndex >= 0 And cmb61.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb61.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb71.SelectedIndex = Me.cmb71.SelectedIndex
                frmIT_5A_G3.cmb71.SelectedIndex = Me.cmb71.SelectedIndex
            End If
        End If
        If cmb72.SelectedIndex >= 0 And cmb62.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb62.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb72.SelectedIndex = Me.cmb72.SelectedIndex
                frmIT_5A_G3.cmb72.SelectedIndex = Me.cmb72.SelectedIndex
            End If
        End If
        If cmb73.SelectedIndex >= 0 And cmb63.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb63.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb73.SelectedIndex = Me.cmb73.SelectedIndex
                frmIT_5A_G3.cmb73.SelectedIndex = Me.cmb73.SelectedIndex
            End If
        End If
        If cmb74.SelectedIndex >= 0 And cmb64.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb64.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb74.SelectedIndex = Me.cmb74.SelectedIndex
                frmIT_5A_G3.cmb74.SelectedIndex = Me.cmb74.SelectedIndex
            End If
        End If
        If cmb75.SelectedIndex >= 0 And cmb65.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb65.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb75.SelectedIndex = Me.cmb75.SelectedIndex
                frmIT_5A_G3.cmb75.SelectedIndex = Me.cmb75.SelectedIndex
            End If
        End If
        If cmb76.SelectedIndex >= 0 And cmb66.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb66.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb76.SelectedIndex = Me.cmb76.SelectedIndex
                frmIT_5A_G3.cmb76.SelectedIndex = Me.cmb76.SelectedIndex
            End If
        End If
        If cmb77.SelectedIndex >= 0 And cmb67.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb67.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb77.SelectedIndex = Me.cmb77.SelectedIndex
                frmIT_5A_G3.cmb77.SelectedIndex = Me.cmb77.SelectedIndex
            End If
        End If
        If cmb78.SelectedIndex >= 0 And cmb68.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb68.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb78.SelectedIndex = Me.cmb78.SelectedIndex
                frmIT_5A_G3.cmb78.SelectedIndex = Me.cmb78.SelectedIndex
            End If
        End If
        If cmb79.SelectedIndex >= 0 And cmb69.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb69.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb79.SelectedIndex = Me.cmb79.SelectedIndex
                frmIT_5A_G3.cmb79.SelectedIndex = Me.cmb79.SelectedIndex
            End If
        End If
        If cmb80.SelectedIndex >= 0 And cmb50.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb50.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb80.SelectedIndex = Me.cmb80.SelectedIndex
                frmIT_5A_G3.cmb80.SelectedIndex = Me.cmb80.SelectedIndex
            End If
        End If
        If cmb81.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb81.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb81.SelectedIndex = Me.cmb81.SelectedIndex
                frmIT_5A_G3.cmb81.SelectedIndex = Me.cmb81.SelectedIndex
            End If
        End If
        If cmb82.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb82.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb82.SelectedIndex = Me.cmb82.SelectedIndex
                frmIT_5A_G3.cmb82.SelectedIndex = Me.cmb82.SelectedIndex
            End If
        End If
        If cmb83.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb83.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb83.SelectedIndex = Me.cmb83.SelectedIndex
                frmIT_5A_G3.cmb83.SelectedIndex = Me.cmb83.SelectedIndex
            End If
        End If
        If cmb84.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb84.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb84.SelectedIndex = Me.cmb84.SelectedIndex
                frmIT_5A_G3.cmb84.SelectedIndex = Me.cmb84.SelectedIndex
            End If
        End If
        If cmb85.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb85.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb85.SelectedIndex = Me.cmb85.SelectedIndex
                frmIT_5A_G3.cmb85.SelectedIndex = Me.cmb85.SelectedIndex
            End If
        End If
        If cmb86.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb86.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb86.SelectedIndex = Me.cmb86.SelectedIndex
                frmIT_5A_G3.cmb86.SelectedIndex = Me.cmb86.SelectedIndex
            End If
        End If
        If cmb87.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb87.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb87.SelectedIndex = Me.cmb87.SelectedIndex
                frmIT_5A_G3.cmb87.SelectedIndex = Me.cmb87.SelectedIndex
            End If
        End If
        If cmb88.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb88.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb88.SelectedIndex = Me.cmb88.SelectedIndex
                frmIT_5A_G3.cmb88.SelectedIndex = Me.cmb88.SelectedIndex
            End If
        End If
        If cmb89.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb89.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb89.SelectedIndex = Me.cmb89.SelectedIndex
                frmIT_5A_G3.cmb89.SelectedIndex = Me.cmb89.SelectedIndex
            End If
        End If
        If cmb90.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb90.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb90.SelectedIndex = Me.cmb90.SelectedIndex
                frmIT_5A_G3.cmb90.SelectedIndex = Me.cmb90.SelectedIndex
            End If
        End If
        If cmb91.SelectedIndex >= 0 And cmb81.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb81.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb91.SelectedIndex = Me.cmb91.SelectedIndex
                frmIT_5A_G3.cmb91.SelectedIndex = Me.cmb91.SelectedIndex
            End If
        End If
        If cmb92.SelectedIndex >= 0 And cmb82.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb82.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb92.SelectedIndex = Me.cmb92.SelectedIndex
                frmIT_5A_G3.cmb92.SelectedIndex = Me.cmb92.SelectedIndex
            End If
        End If
        If cmb93.SelectedIndex >= 0 And cmb83.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb83.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb93.SelectedIndex = Me.cmb93.SelectedIndex
                frmIT_5A_G3.cmb93.SelectedIndex = Me.cmb93.SelectedIndex
            End If
        End If
        If cmb94.SelectedIndex >= 0 And cmb84.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb84.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb94.SelectedIndex = Me.cmb94.SelectedIndex
                frmIT_5A_G3.cmb94.SelectedIndex = Me.cmb94.SelectedIndex
            End If
        End If
        If cmb95.SelectedIndex >= 0 And cmb85.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb85.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb95.SelectedIndex = Me.cmb95.SelectedIndex
                frmIT_5A_G3.cmb95.SelectedIndex = Me.cmb95.SelectedIndex
            End If
        End If
        If cmb96.SelectedIndex >= 0 And cmb86.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb86.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb96.SelectedIndex = Me.cmb96.SelectedIndex
                frmIT_5A_G3.cmb96.SelectedIndex = Me.cmb96.SelectedIndex
            End If
        End If
        If cmb97.SelectedIndex >= 0 And cmb87.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb87.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb97.SelectedIndex = Me.cmb97.SelectedIndex
                frmIT_5A_G3.cmb97.SelectedIndex = Me.cmb97.SelectedIndex
            End If
        End If
        If cmb98.SelectedIndex >= 0 And cmb88.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb88.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb98.SelectedIndex = Me.cmb98.SelectedIndex
                frmIT_5A_G3.cmb98.SelectedIndex = Me.cmb98.SelectedIndex
            End If
        End If
        If cmb99.SelectedIndex >= 0 And cmb89.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb89.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb99.SelectedIndex = Me.cmb99.SelectedIndex
                frmIT_5A_G3.cmb99.SelectedIndex = Me.cmb99.SelectedIndex
            End If
        End If
        If cmb100.SelectedIndex >= 0 And cmb90.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G1(cmb90.SelectedIndex).Type = "L" Then
                frmIT_5A_G1.cmb100.SelectedIndex = Me.cmb100.SelectedIndex
                frmIT_5A_G3.cmb100.SelectedIndex = Me.cmb100.SelectedIndex
            End If
        End If
    End Sub
    Private Sub ToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs)

        ' CheckCorrectness()            // Remove comment to enable correctness check   (Removes incomplete entries)   

        ' Declare Variables
        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook

        ' Start Excel and get Application object.
        oXL = CreateObject("Excel.Application")
        oXL.Visible = True

        Dim Sheet_IT_5A_G2 As Excel.Worksheet
        ' Get new workbooks.

        oWB = oXL.Workbooks.Add
        Sheet_IT_5A_G2 = oWB.ActiveSheet

        ' Format A1:K1 as bold, vertical alignment = center.
        With Sheet_IT_5A_G2.Range("A1", "K1")
            .Font.Bold = True
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End With
        ' Format A1:A6 as bold, vertical alignment = center.
        With Sheet_IT_5A_G2.Range("A1", "A6")
            .Font.Bold = True
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End With

        ' Add table headers going cell by cell.
        Sheet_IT_5A_G2.Cells(1, 1).Value = ""
        Sheet_IT_5A_G2.Cells(1, 2).Value = "8 - 9 AM"
        Sheet_IT_5A_G2.Cells(1, 3).Value = "9 - 10 AM"
        Sheet_IT_5A_G2.Cells(1, 4).Value = "10 - 11 AM"
        Sheet_IT_5A_G2.Cells(1, 5).Value = "11 - 12 AM"
        Sheet_IT_5A_G2.Cells(1, 6).Value = "12 - 1 PM"
        Sheet_IT_5A_G2.Cells(1, 7).Value = "1 - 2 PM"
        Sheet_IT_5A_G2.Cells(1, 8).Value = "2 - 3 PM"
        Sheet_IT_5A_G2.Cells(1, 9).Value = "3 - 4 PM"
        Sheet_IT_5A_G2.Cells(1, 10).Value = "4 - 5 PM"
        Sheet_IT_5A_G2.Cells(1, 11).Value = "5 - 6 PM"
        Sheet_IT_5A_G2.Cells(2, 1).Value = "Monday"
        Sheet_IT_5A_G2.Cells(3, 1).Value = "Tuesday"
        Sheet_IT_5A_G2.Cells(4, 1).Value = "Wednessday"
        Sheet_IT_5A_G2.Cells(5, 1).Value = "Thursday"
        Sheet_IT_5A_G2.Cells(6, 1).Value = "Friday"

        ' Enter Values from ComboBox

        ' Monday

        Sheet_IT_5A_G2.Cells(2, 2).Value = Me.cmb1.Text + Me.cmb11.Text
        Sheet_IT_5A_G2.Cells(2, 3).Value = Me.cmb2.Text + Me.cmb12.Text
        Sheet_IT_5A_G2.Cells(2, 4).Value = Me.cmb3.Text + Me.cmb13.Text
        Sheet_IT_5A_G2.Cells(2, 5).Value = Me.cmb4.Text + Me.cmb14.Text
        Sheet_IT_5A_G2.Cells(2, 6).Value = Me.cmb5.Text + Me.cmb15.Text
        Sheet_IT_5A_G2.Cells(2, 7).Value = Me.cmb6.Text + Me.cmb16.Text
        Sheet_IT_5A_G2.Cells(2, 8).Value = Me.cmb7.Text + Me.cmb17.Text
        Sheet_IT_5A_G2.Cells(2, 9).Value = Me.cmb8.Text + Me.cmb18.Text
        Sheet_IT_5A_G2.Cells(2, 10).Value = Me.cmb9.Text + Me.cmb19.Text
        Sheet_IT_5A_G2.Cells(2, 11).Value = Me.cmb10.Text + Me.cmb20.Text

        ' Tuesday

        Sheet_IT_5A_G2.Cells(3, 2).Value = Me.cmb21.Text + cmb31.Text
        Sheet_IT_5A_G2.Cells(3, 3).Value = Me.cmb22.Text + cmb32.Text
        Sheet_IT_5A_G2.Cells(3, 4).Value = Me.cmb23.Text + cmb33.Text
        Sheet_IT_5A_G2.Cells(3, 5).Value = Me.cmb24.Text + cmb34.Text
        Sheet_IT_5A_G2.Cells(3, 6).Value = Me.cmb25.Text + cmb35.Text
        Sheet_IT_5A_G2.Cells(3, 7).Value = Me.cmb26.Text + cmb36.Text
        Sheet_IT_5A_G2.Cells(3, 8).Value = Me.cmb27.Text + cmb37.Text
        Sheet_IT_5A_G2.Cells(3, 9).Value = Me.cmb28.Text + cmb38.Text
        Sheet_IT_5A_G2.Cells(3, 10).Value = Me.cmb29.Text + cmb39.Text
        Sheet_IT_5A_G2.Cells(3, 11).Value = Me.cmb30.Text + cmb40.Text

        ' Wednessday

        Sheet_IT_5A_G2.Cells(4, 2).Value = Me.cmb41.Text + cmb51.Text
        Sheet_IT_5A_G2.Cells(4, 3).Value = Me.cmb42.Text + cmb52.Text
        Sheet_IT_5A_G2.Cells(4, 4).Value = Me.cmb43.Text + cmb53.Text
        Sheet_IT_5A_G2.Cells(4, 5).Value = Me.cmb44.Text + cmb54.Text
        Sheet_IT_5A_G2.Cells(4, 6).Value = Me.cmb45.Text + cmb55.Text
        Sheet_IT_5A_G2.Cells(4, 7).Value = Me.cmb46.Text + cmb56.Text
        Sheet_IT_5A_G2.Cells(4, 8).Value = Me.cmb47.Text + cmb57.Text
        Sheet_IT_5A_G2.Cells(4, 9).Value = Me.cmb48.Text + cmb58.Text
        Sheet_IT_5A_G2.Cells(4, 10).Value = Me.cmb49.Text + cmb59.Text
        Sheet_IT_5A_G2.Cells(4, 11).Value = Me.cmb50.Text + cmb60.Text

        ' Thursday

        Sheet_IT_5A_G2.Cells(5, 2).Value = Me.cmb61.Text + cmb71.Text
        Sheet_IT_5A_G2.Cells(5, 3).Value = Me.cmb62.Text + cmb72.Text
        Sheet_IT_5A_G2.Cells(5, 4).Value = Me.cmb63.Text + cmb73.Text
        Sheet_IT_5A_G2.Cells(5, 5).Value = Me.cmb64.Text + cmb74.Text
        Sheet_IT_5A_G2.Cells(5, 6).Value = Me.cmb65.Text + cmb75.Text
        Sheet_IT_5A_G2.Cells(5, 7).Value = Me.cmb66.Text + cmb76.Text
        Sheet_IT_5A_G2.Cells(5, 8).Value = Me.cmb67.Text + cmb77.Text
        Sheet_IT_5A_G2.Cells(5, 9).Value = Me.cmb68.Text + cmb78.Text
        Sheet_IT_5A_G2.Cells(5, 10).Value = Me.cmb69.Text + cmb79.Text
        Sheet_IT_5A_G2.Cells(5, 11).Value = Me.cmb70.Text + cmb80.Text

        ' Friday

        Sheet_IT_5A_G2.Cells(6, 2).Value = Me.cmb81.Text + cmb91.Text
        Sheet_IT_5A_G2.Cells(6, 3).Value = Me.cmb82.Text + cmb92.Text
        Sheet_IT_5A_G2.Cells(6, 4).Value = Me.cmb83.Text + cmb93.Text
        Sheet_IT_5A_G2.Cells(6, 5).Value = Me.cmb84.Text + cmb94.Text
        Sheet_IT_5A_G2.Cells(6, 6).Value = Me.cmb85.Text + cmb95.Text
        Sheet_IT_5A_G2.Cells(6, 7).Value = Me.cmb86.Text + cmb96.Text
        Sheet_IT_5A_G2.Cells(6, 8).Value = Me.cmb87.Text + cmb97.Text
        Sheet_IT_5A_G2.Cells(6, 9).Value = Me.cmb88.Text + cmb98.Text
        Sheet_IT_5A_G2.Cells(6, 10).Value = Me.cmb89.Text + cmb99.Text
        Sheet_IT_5A_G2.Cells(6, 11).Value = Me.cmb90.Text + cmb100.Text

    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs)
        PrintForm1.PrintAction = Printing.PrintAction.PrintToPrinter
        PrintDialog1.ShowDialog()     ' Manually set page orientation to Landscape
        PrintForm1.Print()
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Some problems still exist

        PrintForm1.PrintAction = Printing.PrintAction.PrintToPreview
        'PrintDialog1.ShowDialog()
        PrintForm1.Print()
    End Sub

    Private Sub ONToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        ColorFlag = True
        ONToolStripMenuItem1.Checked = True
        OFFToolStripMenuItem1.Checked = False
        RefreshColor()
    End Sub

    Private Sub OFFToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        ColorFlag = False
        ONToolStripMenuItem1.Checked = False
        OFFToolStripMenuItem1.Checked = True
        RefreshColor()
    End Sub


    Private Sub cmb1_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb1.Text = "--none--" Then
            cmb1.SelectedIndex = -1
            RefreshList()
        ElseIf cmb1.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb1.SelectedIndex).Type = "P" Then
                cmb1.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub


    Private Sub cmb2_SelectedIdexChanged(sender As Object, e As EventArgs)
        If cmb2.Text = "--none--" Then
            cmb2.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb3.SelectedIndex).Type = "P" Then
                cmb3.SelectedIndex = -1
                cmb4.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb2.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).Type = "P" Then
                cmb3.SelectedIndex = cmb2.SelectedIndex
                cmb4.SelectedIndex = cmb2.SelectedIndex
            ElseIf cmb3.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb3.SelectedIndex).Type = "P" Then
                    cmb3.SelectedIndex = -1
                    cmb4.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb3_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb3.Text = "--none--" Then
            cmb3.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).Type = "P" Then
                cmb2.SelectedIndex = -1
                cmb4.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb3.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb3.SelectedIndex).Type = "P" Then
                cmb2.SelectedIndex = cmb3.SelectedIndex
                cmb4.SelectedIndex = cmb3.SelectedIndex
            ElseIf cmb2.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).Type = "P" Then
                    cmb2.SelectedIndex = -1
                    cmb4.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb4_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb4.Text = "--none--" Then
            cmb4.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).Type = "P" Then
                cmb2.SelectedIndex = -1
                cmb3.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb4.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb4.SelectedIndex).Type = "P" Then
                cmb3.SelectedIndex = cmb4.SelectedIndex
                cmb2.SelectedIndex = cmb4.SelectedIndex
            ElseIf cmb2.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb2.SelectedIndex).Type = "P" Then
                    cmb2.SelectedIndex = -1
                    cmb3.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb5_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb5.Text = "--none--" Then
            cmb5.SelectedIndex = -1
            RefreshList()
        ElseIf cmb5.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb5.SelectedIndex).Type = "P" Then
                cmb5.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb6_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb6.Text = "--none--" Then
            cmb6.SelectedIndex = -1
            RefreshList()
        ElseIf cmb6.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb6.SelectedIndex).Type = "P" Then
                cmb6.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb7_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb7.Text = "--none--" Then
            cmb7.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb8.SelectedIndex).Type = "P" Then
                cmb8.SelectedIndex = -1
                cmb9.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb7.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb7.SelectedIndex).Type = "P" Then
                cmb8.SelectedIndex = cmb7.SelectedIndex
                cmb9.SelectedIndex = cmb7.SelectedIndex
            ElseIf cmb8.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb8.SelectedIndex).Type = "P" Then
                    cmb8.SelectedIndex = -1
                    cmb9.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb8_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb8.Text = "--none--" Then
            cmb8.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb8.SelectedIndex).Type = "P" Then
                cmb7.SelectedIndex = -1
                cmb9.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb8.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb8.SelectedIndex).Type = "P" Then
                cmb7.SelectedIndex = cmb8.SelectedIndex
                cmb9.SelectedIndex = cmb8.SelectedIndex
            ElseIf cmb7.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb7.SelectedIndex).Type = "P" Then
                    cmb7.SelectedIndex = -1
                    cmb9.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb9_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb9.Text = "--none--" Then
            cmb9.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb7.SelectedIndex).Type = "P" Then
                cmb8.SelectedIndex = -1
                cmb7.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb9.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb9.SelectedIndex).Type = "P" Then
                cmb7.SelectedIndex = cmb9.SelectedIndex
                cmb8.SelectedIndex = cmb9.SelectedIndex
            ElseIf cmb7.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb7.SelectedIndex).Type = "P" Then
                    cmb8.SelectedIndex = -1
                    cmb7.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb10_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb10.Text = "--none--" Then
            cmb10.SelectedIndex = -1
            RefreshList()
        ElseIf cmb10.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb10.SelectedIndex).Type = "P" Then
                cmb10.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb11_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb11.Text = "--none--" Then
            cmb11.SelectedIndex = -1
            RefreshList()
        ElseIf cmb11.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb11.SelectedIndex).Type = "LB" Then
                cmb11.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub


    Private Sub cmb12_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb12.Text = "--none--" Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb13.SelectedIndex).Type = "LB" Then
                cmb13.SelectedIndex = -1
                cmb14.SelectedIndex = -1
            End If
            cmb12.SelectedIndex = -1
            RefreshList()
        ElseIf cmb12.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb12.SelectedIndex).Type = "LB" Then
                cmb13.SelectedIndex = cmb12.SelectedIndex
                cmb14.SelectedIndex = cmb12.SelectedIndex
            ElseIf cmb13.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb13.SelectedIndex).Type = "LB" Then
                    cmb13.SelectedIndex = -1
                    cmb14.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb13_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb13.Text = "--none--" Then
            cmb13.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb12.SelectedIndex).Type = "LB" Then
                cmb12.SelectedIndex = -1
                cmb14.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb13.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb13.SelectedIndex).Type = "LB" Then
                cmb12.SelectedIndex = cmb13.SelectedIndex
                cmb14.SelectedIndex = cmb13.SelectedIndex
            ElseIf cmb14.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb14.SelectedIndex).Type = "LB" Then
                    cmb12.SelectedIndex = -1
                    cmb14.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb14_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb14.Text = "--none--" Then
            cmb14.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb12.SelectedIndex).Type = "LB" Then
                cmb12.SelectedIndex = -1
                cmb13.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb14.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb14.SelectedIndex).Type = "LB" Then
                cmb12.SelectedIndex = cmb14.SelectedIndex
                cmb13.SelectedIndex = cmb14.SelectedIndex
            ElseIf cmb12.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb12.SelectedIndex).Type = "LB" Then
                    cmb12.SelectedIndex = -1
                    cmb13.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb15_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb15.Text = "--none--" Then
            cmb15.SelectedIndex = -1
            RefreshList()
        ElseIf cmb15.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb15.SelectedIndex).Type = "LB" Then
                cmb15.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb16_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb16.Text = "--none--" Then
            cmb16.SelectedIndex = -1
            RefreshList()
        ElseIf cmb16.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb16.SelectedIndex).Type = "LB" Then
                cmb16.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb17_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb17.Text = "--none--" Then
            cmb17.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb18.SelectedIndex).Type = "LB" Then
                cmb18.SelectedIndex = -1
                cmb19.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb17.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb17.SelectedIndex).Type = "LB" Then
                cmb18.SelectedIndex = cmb17.SelectedIndex
                cmb19.SelectedIndex = cmb17.SelectedIndex
            ElseIf cmb19.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb19.SelectedIndex).Type = "LB" Then
                    cmb18.SelectedIndex = -1
                    cmb19.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb18_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb18.Text = "--none--" Then
            cmb18.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb19.SelectedIndex).Type = "LB" Then
                cmb19.SelectedIndex = -1
                cmb17.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb18.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb18.SelectedIndex).Type = "LB" Then
                cmb19.SelectedIndex = cmb18.SelectedIndex
                cmb17.SelectedIndex = cmb18.SelectedIndex
            ElseIf cmb19.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb19.SelectedIndex).Type = "LB" Then
                    cmb19.SelectedIndex = -1
                    cmb17.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb19_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb19.Text = "--none--" Then
            cmb19.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb18.SelectedIndex).Type = "LB" Then
                cmb18.SelectedIndex = -1
                cmb17.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb19.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb19.SelectedIndex).Type = "LB" Then
                cmb18.SelectedIndex = cmb19.SelectedIndex
                cmb17.SelectedIndex = cmb19.SelectedIndex
            ElseIf cmb17.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb17.SelectedIndex).Type = "LB" Then
                    cmb18.SelectedIndex = -1
                    cmb17.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb20_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb20.Text = "--none--" Then
            cmb20.SelectedIndex = -1
            RefreshList()
        ElseIf cmb20.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb20.SelectedIndex).Type = "LB" Then
                cmb20.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb21_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb21.Text = "--none--" Then
            cmb21.SelectedIndex = -1
            RefreshList()
        ElseIf cmb21.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb21.SelectedIndex).Type = "P" Then
                cmb20.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb22_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb22.Text = "--none--" Then
            cmb22.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb23.SelectedIndex).Type = "P" Then
                cmb23.SelectedIndex = -1
                cmb24.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb22.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb22.SelectedIndex).Type = "P" Then
                cmb23.SelectedIndex = cmb22.SelectedIndex
                cmb24.SelectedIndex = cmb22.SelectedIndex
            ElseIf cmb3.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb24.SelectedIndex).Type = "P" Then
                    cmb23.SelectedIndex = -1
                    cmb24.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb23_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb23.Text = "--none--" Then
            cmb23.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb22.SelectedIndex).Type = "P" Then
                cmb22.SelectedIndex = -1
                cmb24.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb23.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb23.SelectedIndex).Type = "P" Then
                cmb22.SelectedIndex = cmb23.SelectedIndex
                cmb24.SelectedIndex = cmb23.SelectedIndex
            ElseIf cmb24.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb24.SelectedIndex).Type = "P" Then
                    cmb22.SelectedIndex = -1
                    cmb24.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb24_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb24.Text = "--none--" Then
            cmb24.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb22.SelectedIndex).Type = "P" Then
                cmb22.SelectedIndex = -1
                cmb23.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb24.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb24.SelectedIndex).Type = "P" Then
                cmb22.SelectedIndex = cmb24.SelectedIndex
                cmb23.SelectedIndex = cmb24.SelectedIndex
            ElseIf cmb22.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb22.SelectedIndex).Type = "P" Then
                    cmb22.SelectedIndex = -1
                    cmb23.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb25_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb25.Text = "--none--" Then
            cmb25.SelectedIndex = -1
            RefreshList()
        ElseIf cmb25.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb25.SelectedIndex).Type = "P" Then
                cmb25.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb26_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb26.Text = "--none--" Then
            cmb26.SelectedIndex = -1
            RefreshList()
        ElseIf cmb26.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb26.SelectedIndex).Type = "P" Then
                cmb26.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb27_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb27.Text = "--none--" Then
            cmb27.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb28.SelectedIndex).Type = "P" Then
                cmb28.SelectedIndex = -1
                cmb29.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb27.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb27.SelectedIndex).Type = "P" Then
                cmb28.SelectedIndex = cmb27.SelectedIndex
                cmb29.SelectedIndex = cmb27.SelectedIndex
            ElseIf cmb29.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb29.SelectedIndex).Type = "P" Then
                    cmb28.SelectedIndex = -1
                    cmb29.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb28_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb28.Text = "--none--" Then
            cmb28.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb27.SelectedIndex).Type = "P" Then
                cmb27.SelectedIndex = -1
                cmb29.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb28.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb28.SelectedIndex).Type = "P" Then
                cmb27.SelectedIndex = cmb28.SelectedIndex
                cmb29.SelectedIndex = cmb28.SelectedIndex
            ElseIf cmb29.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb29.SelectedIndex).Type = "P" Then
                    cmb27.SelectedIndex = -1
                    cmb29.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb29_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb29.Text = "--none--" Then
            cmb29.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb27.SelectedIndex).Type = "P" Then
                cmb27.SelectedIndex = -1
                cmb28.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb29.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb29.SelectedIndex).Type = "P" Then
                cmb27.SelectedIndex = cmb29.SelectedIndex
                cmb28.SelectedIndex = cmb29.SelectedIndex
            ElseIf cmb27.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb27.SelectedIndex).Type = "P" Then
                    cmb27.SelectedIndex = -1
                    cmb28.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb30_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb30.Text = "--none--" Then
            cmb30.SelectedIndex = -1
            RefreshList()
        ElseIf cmb30.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb30.SelectedIndex).Type = "P" Then
                cmb30.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb31_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb31.Text = "--none--" Then
            cmb31.SelectedIndex = -1
            RefreshList()
        ElseIf cmb31.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb31.SelectedIndex).Type = "LB" Then
                cmb31.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb32_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb32.Text = "--none--" Then
            cmb32.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb33.SelectedIndex).Type = "LB" Then
                cmb33.SelectedIndex = -1
                cmb34.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb32.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb32.SelectedIndex).Type = "LB" Then
                cmb33.SelectedIndex = cmb32.SelectedIndex
                cmb34.SelectedIndex = cmb32.SelectedIndex
            ElseIf cmb34.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb34.SelectedIndex).Type = "LB" Then
                    cmb33.SelectedIndex = -1
                    cmb34.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb33_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb33.Text = "--none--" Then
            cmb33.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb32.SelectedIndex).Type = "LB" Then
                cmb32.SelectedIndex = -1
                cmb34.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb33.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb33.SelectedIndex).Type = "LB" Then
                cmb32.SelectedIndex = cmb33.SelectedIndex
                cmb34.SelectedIndex = cmb33.SelectedIndex
            ElseIf cmb34.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb34.SelectedIndex).Type = "LB" Then
                    cmb32.SelectedIndex = -1
                    cmb34.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb34_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb34.Text = "--none--" Then
            cmb34.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb33.SelectedIndex).Type = "LB" Then
                cmb32.SelectedIndex = -1
                cmb33.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb34.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb34.SelectedIndex).Type = "LB" Then
                cmb32.SelectedIndex = cmb34.SelectedIndex
                cmb33.SelectedIndex = cmb34.SelectedIndex
            ElseIf cmb32.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb32.SelectedIndex).Type = "LB" Then
                    cmb32.SelectedIndex = -1
                    cmb33.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb35_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb35.Text = "--none--" Then
            cmb35.SelectedIndex = -1
            RefreshList()
        ElseIf cmb35.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb35.SelectedIndex).Type = "LB" Then
                cmb35.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb36_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb36.Text = "--none--" Then
            cmb36.SelectedIndex = -1
            RefreshList()
        ElseIf cmb36.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb36.SelectedIndex).Type = "LB" Then
                cmb36.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb37_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb37.Text = "--none--" Then
            cmb37.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb38.SelectedIndex).Type = "LB" Then
                cmb38.SelectedIndex = -1
                cmb39.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb37.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb37.SelectedIndex).Type = "LB" Then
                cmb38.SelectedIndex = cmb37.SelectedIndex
                cmb39.SelectedIndex = cmb37.SelectedIndex
            ElseIf cmb39.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb39.SelectedIndex).Type = "LB" Then
                    cmb38.SelectedIndex = -1
                    cmb39.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb38_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb38.Text = "--none--" Then
            cmb38.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb39.SelectedIndex).Type = "LB" Then
                cmb37.SelectedIndex = -1
                cmb39.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb38.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb38.SelectedIndex).Type = "LB" Then
                cmb37.SelectedIndex = cmb38.SelectedIndex
                cmb39.SelectedIndex = cmb38.SelectedIndex
            ElseIf cmb39.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb39.SelectedIndex).Type = "LB" Then
                    cmb37.SelectedIndex = -1
                    cmb39.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb39_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb39.Text = "--none--" Then
            cmb39.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb37.SelectedIndex).Type = "LB" Then
                cmb37.SelectedIndex = -1
                cmb38.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb39.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb39.SelectedIndex).Type = "LB" Then
                cmb37.SelectedIndex = cmb39.SelectedIndex
                cmb38.SelectedIndex = cmb39.SelectedIndex
            ElseIf cmb37.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb37.SelectedIndex).Type = "LB" Then
                    cmb37.SelectedIndex = -1
                    cmb38.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb40_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb40.Text = "--none--" Then
            cmb40.SelectedIndex = -1
            RefreshList()
        ElseIf cmb40.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb40.SelectedIndex).Type = "LB" Then
                cmb40.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb41_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb41.Text = "--none--" Then
            cmb41.SelectedIndex = -1
            RefreshList()
        ElseIf cmb41.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb41.SelectedIndex).Type = "P" Then
                cmb41.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb42_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb42.Text = "--none--" Then
            cmb42.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb43.SelectedIndex).Type = "P" Then
                cmb43.SelectedIndex = -1
                cmb44.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb42.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb42.SelectedIndex).Type = "P" Then
                cmb43.SelectedIndex = cmb42.SelectedIndex
                cmb44.SelectedIndex = cmb42.SelectedIndex
            ElseIf cmb44.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb44.SelectedIndex).Type = "P" Then
                    cmb43.SelectedIndex = -1
                    cmb44.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb43_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb43.Text = "--none--" Then
            cmb43.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb42.SelectedIndex).Type = "P" Then
                cmb42.SelectedIndex = -1
                cmb44.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb43.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb43.SelectedIndex).Type = "P" Then
                cmb42.SelectedIndex = cmb43.SelectedIndex
                cmb44.SelectedIndex = cmb43.SelectedIndex
            ElseIf cmb44.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb44.SelectedIndex).Type = "P" Then
                    cmb42.SelectedIndex = -1
                    cmb44.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb44_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb44.Text = "--none--" Then
            cmb44.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb42.SelectedIndex).Type = "P" Then
                cmb42.SelectedIndex = -1
                cmb43.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb44.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb44.SelectedIndex).Type = "P" Then
                cmb42.SelectedIndex = cmb44.SelectedIndex
                cmb43.SelectedIndex = cmb44.SelectedIndex
            ElseIf cmb42.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb42.SelectedIndex).Type = "P" Then
                    cmb42.SelectedIndex = -1
                    cmb43.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb45_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb45.Text = "--none--" Then
            cmb45.SelectedIndex = -1
            RefreshList()
        ElseIf cmb45.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb45.SelectedIndex).Type = "P" Then
                cmb45.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb46_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb46.Text = "--none--" Then
            cmb46.SelectedIndex = -1
            RefreshList()
        ElseIf cmb46.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb46.SelectedIndex).Type = "P" Then
                cmb46.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb47_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb47.Text = "--none--" Then
            cmb47.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb48.SelectedIndex).Type = "P" Then
                cmb48.SelectedIndex = -1
                cmb49.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb47.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb47.SelectedIndex).Type = "P" Then
                cmb48.SelectedIndex = cmb47.SelectedIndex
                cmb49.SelectedIndex = cmb47.SelectedIndex
            ElseIf cmb49.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb49.SelectedIndex).Type = "P" Then
                    cmb48.SelectedIndex = -1
                    cmb49.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb48_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb48.Text = "--none--" Then
            cmb48.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb47.SelectedIndex).Type = "P" Then
                cmb47.SelectedIndex = -1
                cmb49.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb48.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb48.SelectedIndex).Type = "P" Then
                cmb47.SelectedIndex = cmb48.SelectedIndex
                cmb49.SelectedIndex = cmb48.SelectedIndex
            ElseIf frmMain.DB_oddDataSet1.IT_5A_G2(cmb49.SelectedIndex).Type = "P" Then
                cmb47.SelectedIndex = -1
                cmb49.SelectedIndex = -1
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb49_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb49.Text = "--none--" Then
            cmb49.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb47.SelectedIndex).Type = "P" Then
                cmb47.SelectedIndex = -1
                cmb48.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb49.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb49.SelectedIndex).Type = "P" Then
                cmb47.SelectedIndex = cmb49.SelectedIndex
                cmb48.SelectedIndex = cmb49.SelectedIndex
            ElseIf cmb47.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb47.SelectedIndex).Type = "P" Then
                    cmb47.SelectedIndex = -1
                    cmb48.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb50_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb50.Text = "--none--" Then
            cmb50.SelectedIndex = -1
            RefreshList()
        ElseIf cmb50.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb50.SelectedIndex).Type = "P" Then
                cmb50.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb51_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb51.Text = "--none--" Then
            cmb51.SelectedIndex = -1
            RefreshList()
        ElseIf cmb51.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb51.SelectedIndex).Type = "LB" Then
                cmb51.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb52_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb52.Text = "--none--" Then
            cmb52.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb53.SelectedIndex).Type = "LB" Then
                cmb53.SelectedIndex = -1
                cmb54.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb52.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb52.SelectedIndex).Type = "LB" Then
                cmb53.SelectedIndex = cmb52.SelectedIndex
                cmb54.SelectedIndex = cmb52.SelectedIndex
            ElseIf cmb54.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb54.SelectedIndex).Type = "LB" Then
                    cmb53.SelectedIndex = -1
                    cmb54.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb53_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb53.Text = "--none--" Then
            cmb53.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb53.SelectedIndex).Type = "LB" Then
                cmb52.SelectedIndex = -1
                cmb54.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb53.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb53.SelectedIndex).Type = "LB" Then
                cmb52.SelectedIndex = cmb53.SelectedIndex
                cmb54.SelectedIndex = cmb53.SelectedIndex
            ElseIf cmb54.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb54.SelectedIndex).Type = "LB" Then
                    cmb52.SelectedIndex = -1
                    cmb54.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb54_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb54.Text = "--none--" Then
            cmb54.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb52.SelectedIndex).Type = "LB" Then
                cmb52.SelectedIndex = -1
                cmb53.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb54.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb54.SelectedIndex).Type = "LB" Then
                cmb52.SelectedIndex = cmb54.SelectedIndex
                cmb53.SelectedIndex = cmb54.SelectedIndex
            ElseIf cmb52.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb52.SelectedIndex).Type = "LB" Then
                    cmb52.SelectedIndex = -1
                    cmb53.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb55_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb55.Text = "--none--" Then
            cmb55.SelectedIndex = -1
            RefreshList()
        ElseIf cmb55.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb55.SelectedIndex).Type = "LB" Then
                cmb55.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb56_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb56.Text = "--none--" Then
            cmb56.SelectedIndex = -1
            RefreshList()
        ElseIf cmb56.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb56.SelectedIndex).Type = "LB" Then
                cmb56.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb57_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb57.Text = "--none--" Then
            cmb57.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb58.SelectedIndex).Type = "LB" Then
                cmb58.SelectedIndex = -1
                cmb59.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb57.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb57.SelectedIndex).Type = "LB" Then
                cmb58.SelectedIndex = cmb57.SelectedIndex
                cmb59.SelectedIndex = cmb57.SelectedIndex
            ElseIf cmb59.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb59.SelectedIndex).Type = "LB" Then
                    cmb58.SelectedIndex = -1
                    cmb59.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb58_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb58.Text = "--none--" Then
            cmb58.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb57.SelectedIndex).Type = "LB" Then
                cmb57.SelectedIndex = -1
                cmb59.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb58.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb58.SelectedIndex).Type = "LB" Then
                cmb57.SelectedIndex = cmb58.SelectedIndex
                cmb59.SelectedIndex = cmb58.SelectedIndex
            ElseIf cmb59.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb59.SelectedIndex).Type = "LB" Then
                    cmb57.SelectedIndex = -1
                    cmb59.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb59_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb59.Text = "--none--" Then
            cmb59.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb57.SelectedIndex).Type = "LB" Then
                cmb57.SelectedIndex = -1
                cmb58.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb59.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb59.SelectedIndex).Type = "LB" Then
                cmb57.SelectedIndex = cmb59.SelectedIndex
                cmb58.SelectedIndex = cmb59.SelectedIndex
            ElseIf cmb57.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb57.SelectedIndex).Type = "LB" Then
                    cmb57.SelectedIndex = -1
                    cmb58.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb60_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb60.Text = "--none--" Then
            cmb60.SelectedIndex = -1
            RefreshList()
        ElseIf cmb60.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb60.SelectedIndex).Type = "LB" Then
                cmb60.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb61_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb61.Text = "--none--" Then
            cmb61.SelectedIndex = -1
            RefreshList()
        ElseIf cmb61.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb61.SelectedIndex).Type = "P" Then
                cmb61.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb62_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb62.Text = "--none--" Then
            cmb62.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb63.SelectedIndex).Type = "P" Then
                cmb63.SelectedIndex = -1
                cmb64.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb62.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb62.SelectedIndex).Type = "P" Then
                cmb63.SelectedIndex = cmb62.SelectedIndex
                cmb64.SelectedIndex = cmb62.SelectedIndex
            ElseIf cmb64.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb64.SelectedIndex).Type = "P" Then
                    cmb63.SelectedIndex = -1
                    cmb64.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb63_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb63.Text = "--none--" Then
            cmb63.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb62.SelectedIndex).Type = "P" Then
                cmb62.SelectedIndex = -1
                cmb64.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb63.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb63.SelectedIndex).Type = "P" Then
                cmb62.SelectedIndex = cmb63.SelectedIndex
                cmb64.SelectedIndex = cmb63.SelectedIndex
            ElseIf cmb64.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb64.SelectedIndex).Type = "P" Then
                    cmb62.SelectedIndex = -1
                    cmb64.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb64_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb64.Text = "--none--" Then
            cmb64.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb62.SelectedIndex).Type = "P" Then
                cmb62.SelectedIndex = -1
                cmb63.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb64.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb64.SelectedIndex).Type = "P" Then
                cmb62.SelectedIndex = cmb64.SelectedIndex
                cmb63.SelectedIndex = cmb64.SelectedIndex
            ElseIf frmMain.DB_oddDataSet1.IT_5A_G2(cmb62.SelectedIndex).Type = "P" Then
                cmb62.SelectedIndex = -1
                cmb63.SelectedIndex = -1
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb65_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb65.Text = "--none--" Then
            cmb65.SelectedIndex = -1
            RefreshList()
        ElseIf cmb65.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb65.SelectedIndex).Type = "P" Then
                cmb65.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb66_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb66.Text = "--none--" Then
            cmb66.SelectedIndex = -1
            RefreshList()
        ElseIf cmb66.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb66.SelectedIndex).Type = "P" Then
                cmb66.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb67_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb67.Text = "--none--" Then
            cmb67.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb68.SelectedIndex).Type = "P" Then
                cmb68.SelectedIndex = -1
                cmb69.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb67.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb67.SelectedIndex).Type = "P" Then
                cmb68.SelectedIndex = cmb67.SelectedIndex
                cmb69.SelectedIndex = cmb67.SelectedIndex
            ElseIf cmb69.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb69.SelectedIndex).Type = "P" Then
                    cmb68.SelectedIndex = -1
                    cmb69.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb68_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb68.Text = "--none--" Then
            cmb68.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb67.SelectedIndex).Type = "P" Then
                cmb68.SelectedIndex = -1
                cmb69.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb68.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb68.SelectedIndex).Type = "P" Then
                cmb68.SelectedIndex = cmb68.SelectedIndex
                cmb69.SelectedIndex = cmb68.SelectedIndex
            ElseIf cmb69.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb69.SelectedIndex).Type = "P" Then
                    cmb68.SelectedIndex = -1
                    cmb69.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb69_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb69.Text = "--none--" Then
            cmb69.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb67.SelectedIndex).Type = "P" Then
                cmb67.SelectedIndex = -1
                cmb68.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb69.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb69.SelectedIndex).Type = "P" Then
                cmb67.SelectedIndex = cmb69.SelectedIndex
                cmb68.SelectedIndex = cmb69.SelectedIndex
            ElseIf cmb67.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb67.SelectedIndex).Type = "P" Then
                    cmb67.SelectedIndex = -1
                    cmb68.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb70_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb70.Text = "--none--" Then
            cmb70.SelectedIndex = -1
            RefreshList()
        ElseIf cmb70.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb70.SelectedIndex).Type = "P" Then
                cmb70.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb71_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb71.Text = "--none--" Then
            cmb71.SelectedIndex = -1
            RefreshList()
        ElseIf cmb71.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb71.SelectedIndex).Type = "LB" Then
                cmb71.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb72_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb72.Text = "--none--" Then
            cmb72.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb73.SelectedIndex).Type = "LB" Then
                cmb73.SelectedIndex = -1
                cmb74.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb72.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb72.SelectedIndex).Type = "LB" Then
                cmb73.SelectedIndex = cmb72.SelectedIndex
                cmb74.SelectedIndex = cmb72.SelectedIndex
            ElseIf cmb74.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb74.SelectedIndex).Type = "LB" Then
                    cmb73.SelectedIndex = -1
                    cmb74.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb73_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb73.Text = "--none--" Then
            cmb73.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb72.SelectedIndex).Type = "LB" Then
                cmb72.SelectedIndex = -1
                cmb74.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb73.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb73.SelectedIndex).Type = "LB" Then
                cmb73.SelectedIndex = cmb73.SelectedIndex
                cmb74.SelectedIndex = cmb73.SelectedIndex
            ElseIf cmb74.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb74.SelectedIndex).Type = "LB" Then
                    cmb72.SelectedIndex = -1
                    cmb74.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb74_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb74.Text = "--none--" Then
            cmb74.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb72.SelectedIndex).Type = "LB" Then
                cmb73.SelectedIndex = -1
                cmb72.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb74.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb74.SelectedIndex).Type = "LB" Then
                cmb72.SelectedIndex = cmb74.SelectedIndex
                cmb73.SelectedIndex = cmb74.SelectedIndex
            ElseIf cmb72.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb72.SelectedIndex).Type = "LB" Then
                    cmb73.SelectedIndex = -1
                    cmb72.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb75_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb75.Text = "--none--" Then
            cmb75.SelectedIndex = -1
            RefreshList()
        ElseIf cmb75.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb75.SelectedIndex).Type = "LB" Then
                cmb75.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb76_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb76.Text = "--none--" Then
            cmb76.SelectedIndex = -1
            RefreshList()
        ElseIf cmb76.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb76.SelectedIndex).Type = "LB" Then
                cmb76.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb77_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb77.Text = "--none--" Then
            cmb77.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb78.SelectedIndex).Type = "LB" Then
                cmb78.SelectedIndex = -1
                cmb79.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb77.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb77.SelectedIndex).Type = "LB" Then
                cmb78.SelectedIndex = cmb77.SelectedIndex
                cmb79.SelectedIndex = cmb77.SelectedIndex
            ElseIf cmb79.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb79.SelectedIndex).Type = "LB" Then
                    cmb78.SelectedIndex = -1
                    cmb79.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb78_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb78.Text = "--none--" Then
            cmb78.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb77.SelectedIndex).Type = "LB" Then
                cmb77.SelectedIndex = -1
                cmb79.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb78.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb78.SelectedIndex).Type = "LB" Then
                cmb77.SelectedIndex = cmb78.SelectedIndex
                cmb79.SelectedIndex = cmb78.SelectedIndex
            ElseIf cmb79.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb79.SelectedIndex).Type = "LB" Then
                    cmb77.SelectedIndex = -1
                    cmb79.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb79_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb79.Text = "--none--" Then
            cmb79.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb77.SelectedIndex).Type = "LB" Then
                cmb77.SelectedIndex = -1
                cmb78.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb79.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb79.SelectedIndex).Type = "LB" Then
                cmb77.SelectedIndex = cmb79.SelectedIndex
                cmb78.SelectedIndex = cmb79.SelectedIndex
            ElseIf cmb77.Text <> "" Then
                If frmMain.DB_room_ODDDataSet1.Rooms(cmb77.SelectedIndex).Type = "LB" Then
                    cmb77.SelectedIndex = -1
                    cmb78.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb80_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb80.Text = "--none--" Then
            cmb80.SelectedIndex = -1
            RefreshList()
        ElseIf cmb80.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb80.SelectedIndex).Type = "LB" Then
                cmb80.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub

    Private Sub cmb81_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb81.Text = "--none--" Then
            cmb81.SelectedIndex = -1
            RefreshList()
        ElseIf cmb81.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb81.SelectedIndex).Type = "P" Then
                cmb81.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb82_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb82.Text = "--none--" Then
            cmb82.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb83.SelectedIndex).Type = "P" Then
                cmb83.SelectedIndex = -1
                cmb84.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb82.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb82.SelectedIndex).Type = "P" Then
                cmb83.SelectedIndex = cmb82.SelectedIndex
                cmb84.SelectedIndex = cmb82.SelectedIndex
            ElseIf cmb84.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb84.SelectedIndex).Type = "P" Then
                    cmb83.SelectedIndex = -1
                    cmb84.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb83_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb83.Text = "--none--" Then
            cmb83.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb82.SelectedIndex).Type = "P" Then
                cmb83.SelectedIndex = -1
                cmb84.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb83.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb83.SelectedIndex).Type = "P" Then
                cmb82.SelectedIndex = cmb83.SelectedIndex
                cmb84.SelectedIndex = cmb83.SelectedIndex
            ElseIf cmb84.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb84.SelectedIndex).Type = "P" Then
                    cmb83.SelectedIndex = -1
                    cmb84.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb84_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb84.Text = "--none--" Then
            cmb84.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb82.SelectedIndex).Type = "P" Then
                cmb82.SelectedIndex = -1
                cmb83.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb84.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb84.SelectedIndex).Type = "P" Then
                cmb82.SelectedIndex = cmb84.SelectedIndex
                cmb83.SelectedIndex = cmb84.SelectedIndex
            ElseIf cmb82.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb82.SelectedIndex).Type = "P" Then
                    cmb82.SelectedIndex = -1
                    cmb83.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb85_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb85.Text = "--none--" Then
            cmb85.SelectedIndex = -1
            RefreshList()
        ElseIf cmb85.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb85.SelectedIndex).Type = "P" Then
                cmb85.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb86_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb86.Text = "--none--" Then
            cmb86.SelectedIndex = -1
            RefreshList()
        ElseIf cmb86.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb86.SelectedIndex).Type = "P" Then
                cmb86.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb87_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb87.Text = "--none--" Then
            cmb87.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb88.SelectedIndex).Type = "P" Then
                cmb88.SelectedIndex = -1
                cmb89.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb87.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb87.SelectedIndex).Type = "P" Then
                cmb88.SelectedIndex = cmb87.SelectedIndex
                cmb89.SelectedIndex = cmb87.SelectedIndex
            ElseIf cmb89.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb89.SelectedIndex).Type = "P" Then
                    cmb88.SelectedIndex = -1
                    cmb89.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb88_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb88.Text = "--none--" Then
            cmb88.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb87.SelectedIndex).Type = "P" Then
                cmb87.SelectedIndex = -1
                cmb89.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb88.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb88.SelectedIndex).Type = "P" Then
                cmb87.SelectedIndex = cmb88.SelectedIndex
                cmb89.SelectedIndex = cmb88.SelectedIndex
            ElseIf cmb89.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb89.SelectedIndex).Type = "P" Then
                    cmb87.SelectedIndex = -1
                    cmb89.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb89_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb89.Text = "--none--" Then
            cmb89.SelectedIndex = -1
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb87.SelectedIndex).Type = "P" Then
                cmb87.SelectedIndex = -1
                cmb88.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb89.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb89.SelectedIndex).Type = "P" Then
                cmb87.SelectedIndex = cmb89.SelectedIndex
                cmb88.SelectedIndex = cmb89.SelectedIndex
            ElseIf cmb87.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb87.SelectedIndex).Type = "P" Then
                    cmb87.SelectedIndex = -1
                    cmb88.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb90_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb90.Text = "--none--" Then
            cmb90.SelectedIndex = -1
            RefreshList()
        ElseIf cmb90.SelectedIndex >= 0 Then
            If frmMain.DB_oddDataSet1.IT_5A_G2(cmb90.SelectedIndex).Type = "P" Then
                cmb90.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb91_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb91.Text = "--none--" Then
            cmb91.SelectedIndex = -1
            RefreshList()
        ElseIf cmb91.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb91.SelectedIndex).Type = "LB" Then
                cmb91.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb92_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb92.Text = "--none--" Then
            cmb92.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb93.SelectedIndex).Type = "LB" Then
                cmb93.SelectedIndex = -1
                cmb94.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb92.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb92.SelectedIndex).Type = "LB" Then
                cmb93.SelectedIndex = cmb92.SelectedIndex
                cmb94.SelectedIndex = cmb92.SelectedIndex
            ElseIf cmb94.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb94.SelectedIndex).Type = "P" Then
                    cmb93.SelectedIndex = -1
                    cmb94.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb93_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb93.Text = "--none--" Then
            cmb93.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb94.SelectedIndex).Type = "LB" Then
                cmb94.SelectedIndex = -1
                cmb92.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb93.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb93.SelectedIndex).Type = "LB" Then
                cmb92.SelectedIndex = cmb93.SelectedIndex
                cmb94.SelectedIndex = cmb93.SelectedIndex
            ElseIf cmb94.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb94.SelectedIndex).Type = "P" Then
                    cmb94.SelectedIndex = -1
                    cmb92.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb94_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb94.Text = "--none--" Then
            cmb94.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb94.SelectedIndex).Type = "LB" Then
                cmb92.SelectedIndex = -1
                cmb93.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb94.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb94.SelectedIndex).Type = "LB" Then
                cmb92.SelectedIndex = cmb94.SelectedIndex
                cmb93.SelectedIndex = cmb94.SelectedIndex
            ElseIf cmb92.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb92.SelectedIndex).Type = "P" Then
                    cmb92.SelectedIndex = -1
                    cmb93.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb95_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb95.Text = "--none--" Then
            cmb95.SelectedIndex = -1
            RefreshList()
        ElseIf cmb95.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb95.SelectedIndex).Type = "LB" Then
                cmb95.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb96_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb96.Text = "--none--" Then
            cmb96.SelectedIndex = -1
            RefreshList()
        ElseIf cmb96.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb96.SelectedIndex).Type = "LB" Then
                cmb96.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb97_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb97.Text = "--none--" Then
            cmb97.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb98.SelectedIndex).Type = "LB" Then
                cmb98.SelectedIndex = -1
                cmb99.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb97.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb97.SelectedIndex).Type = "LB" Then
                cmb98.SelectedIndex = cmb97.SelectedIndex
                cmb99.SelectedIndex = cmb97.SelectedIndex
            ElseIf cmb99.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb99.SelectedIndex).Type = "P" Then
                    cmb98.SelectedIndex = -1
                    cmb99.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb98_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb98.Text = "--none--" Then
            cmb98.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb97.SelectedIndex).Type = "LB" Then
                cmb97.SelectedIndex = -1
                cmb99.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb98.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb98.SelectedIndex).Type = "LB" Then
                cmb97.SelectedIndex = cmb98.SelectedIndex
                cmb99.SelectedIndex = cmb98.SelectedIndex
            ElseIf cmb99.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb99.SelectedIndex).Type = "P" Then
                    cmb97.SelectedIndex = -1
                    cmb99.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb99_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb99.Text = "--none--" Then
            cmb99.SelectedIndex = -1
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb97.SelectedIndex).Type = "LB" Then
                cmb97.SelectedIndex = -1
                cmb98.SelectedIndex = -1
            End If
            RefreshList()
        ElseIf cmb99.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb99.SelectedIndex).Type = "LB" Then
                cmb97.SelectedIndex = cmb99.SelectedIndex
                cmb98.SelectedIndex = cmb99.SelectedIndex
            ElseIf cmb97.Text <> "" Then
                If frmMain.DB_oddDataSet1.IT_5A_G2(cmb97.SelectedIndex).Type = "P" Then
                    cmb97.SelectedIndex = -1
                    cmb98.SelectedIndex = -1
                End If
            End If
            RefreshList()
        End If
    End Sub
    Private Sub cmb100_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmb100.Text = "--none--" Then
            cmb100.SelectedIndex = -1
            RefreshList()
        ElseIf cmb100.SelectedIndex >= 0 Then
            If frmMain.DB_room_ODDDataSet1.Rooms(cmb100.SelectedIndex).Type = "LB" Then
                cmb100.SelectedIndex = -1
                MsgBox("Practicals not allowed here..!!!", vbInformation, "Not Allowed !")
            End If
            RefreshList()
        End If
    End Sub

    Private Sub TeachersTimeTableToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Not Implemented Yet
        MsgBox("Sorry! This feature is not available yet...!!!")
    End Sub

    Private Sub MainFormToolStripMenuItem_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmMain.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_3A_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_3A_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_3A_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem6_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_3B_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem6_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_3B_g2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem6_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_3B_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem3_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_4A_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem3_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_4A_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem3_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_4A_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem9_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_4B_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem9_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_4B_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem9_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_4B_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_5A_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Group3ToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_5A_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem7_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_5B_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem7_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_5B_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem7_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_5B_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem4_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_6A_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem4_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
            frmIT_6A_G2.Show()
        Next frm
    End Sub

    Private Sub Group3ToolStripMenuItem4_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_6A_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem10_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_6B_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem10_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_6B_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem10_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_6B_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_7A_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_7A_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_7A_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem8_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_7B_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem8_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_7B_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem8_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_7B_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem5_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_8A_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem5_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_8A_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem5_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_8A_G3.Show()
    End Sub

    Private Sub Group1ToolStripMenuItem11_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_8B_G1.Show()
    End Sub

    Private Sub Group2ToolStripMenuItem11_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_8B_G2.Show()
    End Sub

    Private Sub Group3ToolStripMenuItem11_Click(sender As Object, e As EventArgs)
        For Each frm As Form In Application.OpenForms
            frm.Hide()
        Next frm
        frmIT_8B_G3.Show()
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Shows warning and resets database ie complete rollback
        If MsgBox("All data will be lost..  Be careful...!!!\n Do you still want to continue ???", vbYesNo, "Warning...!!!") = 6 Then
            frmMain.initDB()
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs)
        If MsgBox("Your work will be lost after this!!!  Dont Forget to Save first! Do you still want to continue ???", vbYesNo Or vbCritical, "Warning...!!!") = 6 Then
            End
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' This feature is not implemented yet! 
        MsgBox("This feature has not been implemented yet!", vbInformation, "Unavailable Feature!!!")
    End Sub

    Private Sub CheckCorrectness(sender As Object, e As EventArgs) Handles CheckCorrectnessToolStripMenuItem.Click
        If MsgBox("This might delete some incomplete information on the form!!! Do you still want to continue???", vbYesNo, "Prompt") = 6 Then
            CheckCorrectness()
        End If
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs)
        MsgBox("This feature has not been implemented yet!!!")
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub
End Class
