Imports Microsoft.Office.Interop

Public Class frmMain
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.IT_4A_G1TableAdapter1.Fill(Me.DB_evenDataSet1.IT_4A_G1)
        Me.IT_4A_G2TableAdapter1.Fill(Me.DB_evenDataSet1.IT_4A_G2)
        Me.IT_4A_G3TableAdapter1.Fill(Me.DB_evenDataSet1.IT_4A_G3)

        Me.IT_4B_G1TableAdapter1.Fill(Me.DB_evenDataSet1.IT_4B_G1)
        Me.IT_4B_G2TableAdapter1.Fill(Me.DB_evenDataSet1.IT_4B_G2)
        Me.IT_4B_G3TableAdapter1.Fill(Me.DB_evenDataSet1.IT_4B_G3)

        Me.IT_6A_G1TableAdapter1.Fill(Me.DB_evenDataSet1.IT_6A_G1)
        Me.IT_6A_G2TableAdapter1.Fill(Me.DB_evenDataSet1.IT_6A_G2)
        Me.IT_6A_G3TableAdapter1.Fill(Me.DB_evenDataSet1.IT_6A_G3)

        Me.IT_6B_G1TableAdapter1.Fill(Me.DB_evenDataSet1.IT_6B_G1)
        Me.IT_6B_G2TableAdapter1.Fill(Me.DB_evenDataSet1.IT_6B_G2)
        Me.IT_6B_G3TableAdapter1.Fill(Me.DB_evenDataSet1.IT_6B_G3)

        Me.IT_8A_G1TableAdapter1.Fill(Me.DB_evenDataSet1.IT_8A_G1)
        Me.IT_8A_G2TableAdapter1.Fill(Me.DB_evenDataSet1.IT_8A_G2)
        Me.IT_8A_G3TableAdapter1.Fill(Me.DB_evenDataSet1.IT_8A_G3)

        Me.IT_8B_G1TableAdapter1.Fill(Me.DB_evenDataSet1.IT_8B_G1)
        Me.IT_8B_G2TableAdapter1.Fill(Me.DB_evenDataSet1.IT_8B_G2)
        Me.IT_8B_G3TableAdapter1.Fill(Me.DB_evenDataSet1.IT_8B_G3)


        ' Filling Odd Dataset

        Me.IT_3A_G1TableAdapter1.Fill(Me.DB_oddDataSet1.IT_3A_G1)
        Me.IT_3A_G2TableAdapter1.Fill(Me.DB_oddDataSet1.IT_3A_G2)
        Me.IT_3A_G3TableAdapter1.Fill(Me.DB_oddDataSet1.IT_3A_G3)

        Me.IT_3B_G1TableAdapter1.Fill(Me.DB_oddDataSet1.IT_3B_G1)
        Me.IT_3B_G2TableAdapter1.Fill(Me.DB_oddDataSet1.IT_3B_G2)
        Me.IT_3B_G3TableAdapter1.Fill(Me.DB_oddDataSet1.IT_3B_G3)

        Me.IT_5A_G1TableAdapter1.Fill(Me.DB_oddDataSet1.IT_5A_G1)
        Me.IT_5A_G2TableAdapter1.Fill(Me.DB_oddDataSet1.IT_5A_G2)
        Me.IT_5A_G3TableAdapter1.Fill(Me.DB_oddDataSet1.IT_5A_G3)

        Me.IT_5B_G1TableAdapter1.Fill(Me.DB_oddDataSet1.IT_5B_G1)
        Me.IT_5B_G2TableAdapter1.Fill(Me.DB_oddDataSet1.IT_5B_G2)
        Me.IT_5B_G3TableAdapter1.Fill(Me.DB_oddDataSet1.IT_5B_G3)

        Me.IT_7A_G1TableAdapter1.Fill(Me.DB_oddDataSet1.IT_7A_G1)
        Me.IT_7A_G2TableAdapter1.Fill(Me.DB_oddDataSet1.IT_7A_G2)
        Me.IT_7A_G3TableAdapter1.Fill(Me.DB_oddDataSet1.IT_7A_G3)

        Me.IT_7B_G1TableAdapter1.Fill(Me.DB_oddDataSet1.IT_7B_G1)
        Me.IT_7B_G2TableAdapter1.Fill(Me.DB_oddDataSet1.IT_7B_G2)
        Me.IT_7B_G3TableAdapter1.Fill(Me.DB_oddDataSet1.IT_7B_G3)

        Me.RoomsTableAdapter3.Fill(Me.DB_room_EVENDataSet1.Rooms)
        Me.RoomsTableAdapter2.Fill(Me.DB_room_ODDDataSet1.Rooms)
        Me.TeachersTableAdapter1.Fill(Me.DB_teacher_EVENDataSet1.Teachers)
        Me.TeachersTableAdapter2.Fill(Me.DB_teacher_ODDDataSet1.Teachers)

    End Sub

    Public Sub initDB()

        ' Semester 3                           

        For a = 0 To DB_oddDataSet1.IT_3A_G1.Count - 1

            If DB_oddDataSet1.IT_3A_G1(a).Type = "L" Or DB_oddDataSet1.IT_3A_G1(a).Type = "P" Then

                DB_oddDataSet1.IT_3A_G1(a).No_per_week = 3
                DB_oddDataSet1.IT_3A_G2(a).No_per_week = 3
                DB_oddDataSet1.IT_3A_G3(a).No_per_week = 3
                DB_oddDataSet1.IT_3B_G1(a).No_per_week = 3
                DB_oddDataSet1.IT_3B_G2(a).No_per_week = 3
                DB_oddDataSet1.IT_3B_G3(a).No_per_week = 3

            ElseIf DB_oddDataSet1.IT_3B_G1(a).Type = "T" Then

                DB_oddDataSet1.IT_3A_G1(a).No_per_week = 2
                DB_oddDataSet1.IT_3A_G2(a).No_per_week = 2
                DB_oddDataSet1.IT_3A_G3(a).No_per_week = 2
                DB_oddDataSet1.IT_3B_G1(a).No_per_week = 2
                DB_oddDataSet1.IT_3B_G2(a).No_per_week = 2
                DB_oddDataSet1.IT_3B_G3(a).No_per_week = 2

            End If

            DB_oddDataSet1.IT_3A_G1(a).Taken = 0
            DB_oddDataSet1.IT_3A_G2(a).Taken = 0
            DB_oddDataSet1.IT_3A_G3(a).Taken = 0
            DB_oddDataSet1.IT_3B_G1(a).Taken = 0
            DB_oddDataSet1.IT_3B_G2(a).Taken = 0
            DB_oddDataSet1.IT_3B_G3(a).Taken = 0

        Next a

        'Semester 4

        For a = 0 To DB_evenDataSet1.IT_4A_G1.Count - 1

            If DB_evenDataSet1.IT_4A_G1(a).Type = "L" Or DB_evenDataSet1.IT_4A_G1(a).Type = "P" Then

                DB_evenDataSet1.IT_4A_G1(a).No_per_week = 3
                DB_evenDataSet1.IT_4A_G2(a).No_per_week = 3
                DB_evenDataSet1.IT_4A_G3(a).No_per_week = 3
                DB_evenDataSet1.IT_4B_G1(a).No_per_week = 3
                DB_evenDataSet1.IT_4B_G2(a).No_per_week = 3
                DB_evenDataSet1.IT_4B_G3(a).No_per_week = 3

            ElseIf DB_evenDataSet1.IT_4B_G1(a).Type = "T" Then

                DB_evenDataSet1.IT_4A_G1(a).No_per_week = 2
                DB_evenDataSet1.IT_4A_G2(a).No_per_week = 2
                DB_evenDataSet1.IT_4A_G3(a).No_per_week = 2
                DB_evenDataSet1.IT_4B_G1(a).No_per_week = 2
                DB_evenDataSet1.IT_4B_G2(a).No_per_week = 2
                DB_evenDataSet1.IT_4B_G3(a).No_per_week = 2

            End If

            DB_evenDataSet1.IT_4A_G1(a).Taken = 0
            DB_evenDataSet1.IT_4A_G2(a).Taken = 0
            DB_evenDataSet1.IT_4A_G3(a).Taken = 0
            DB_evenDataSet1.IT_4B_G1(a).Taken = 0
            DB_evenDataSet1.IT_4B_G2(a).Taken = 0
            DB_evenDataSet1.IT_4B_G3(a).Taken = 0

        Next a

        ' Semester 5

        For a = 0 To DB_oddDataSet1.IT_5A_G1.Count - 1

            If DB_oddDataSet1.IT_5A_G1(a).Type = "L" Or DB_oddDataSet1.IT_5A_G1(a).Type = "P" Then

                DB_oddDataSet1.IT_5A_G1(a).No_per_week = 3
                DB_oddDataSet1.IT_5A_G2(a).No_per_week = 3
                DB_oddDataSet1.IT_5A_G3(a).No_per_week = 3
                DB_oddDataSet1.IT_5B_G1(a).No_per_week = 3
                DB_oddDataSet1.IT_5B_G2(a).No_per_week = 3
                DB_oddDataSet1.IT_5B_G3(a).No_per_week = 3

            ElseIf DB_oddDataSet1.IT_5B_G1(a).Type = "T" Then

                DB_oddDataSet1.IT_5A_G1(a).No_per_week = 2
                DB_oddDataSet1.IT_5A_G2(a).No_per_week = 2
                DB_oddDataSet1.IT_5A_G3(a).No_per_week = 2
                DB_oddDataSet1.IT_5B_G1(a).No_per_week = 2
                DB_oddDataSet1.IT_5B_G2(a).No_per_week = 2
                DB_oddDataSet1.IT_5B_G3(a).No_per_week = 2

            End If

            DB_oddDataSet1.IT_5A_G1(a).Taken = 0
            DB_oddDataSet1.IT_5A_G2(a).Taken = 0
            DB_oddDataSet1.IT_5A_G3(a).Taken = 0
            DB_oddDataSet1.IT_5B_G1(a).Taken = 0
            DB_oddDataSet1.IT_5B_G2(a).Taken = 0
            DB_oddDataSet1.IT_5B_G3(a).Taken = 0

        Next a

        ' Semester 6

        For a = 0 To DB_evenDataSet1.IT_6A_G1.Count - 1

            If DB_evenDataSet1.IT_6A_G1(a).Type = "L" Or DB_evenDataSet1.IT_6A_G1(a).Type = "P" Then

                DB_evenDataSet1.IT_6A_G1(a).No_per_week = 3
                DB_evenDataSet1.IT_6A_G2(a).No_per_week = 3
                DB_evenDataSet1.IT_6A_G3(a).No_per_week = 3
                DB_evenDataSet1.IT_6B_G1(a).No_per_week = 3
                DB_evenDataSet1.IT_6B_G2(a).No_per_week = 3
                DB_evenDataSet1.IT_6B_G3(a).No_per_week = 3

            ElseIf DB_evenDataSet1.IT_6B_G1(a).Type = "T" Then

                DB_evenDataSet1.IT_6A_G1(a).No_per_week = 2
                DB_evenDataSet1.IT_6A_G2(a).No_per_week = 2
                DB_evenDataSet1.IT_6A_G3(a).No_per_week = 2
                DB_evenDataSet1.IT_6B_G1(a).No_per_week = 2
                DB_evenDataSet1.IT_6B_G2(a).No_per_week = 2
                DB_evenDataSet1.IT_6B_G3(a).No_per_week = 2

            End If

            DB_evenDataSet1.IT_6A_G1(a).Taken = 0
            DB_evenDataSet1.IT_6A_G2(a).Taken = 0
            DB_evenDataSet1.IT_6A_G3(a).Taken = 0
            DB_evenDataSet1.IT_6B_G1(a).Taken = 0
            DB_evenDataSet1.IT_6B_G2(a).Taken = 0
            DB_evenDataSet1.IT_6B_G3(a).Taken = 0

        Next a

        ' Semester 7

        For a = 0 To DB_oddDataSet1.IT_7A_G1.Count - 1

            If DB_oddDataSet1.IT_7A_G1(a).Type = "L" Or DB_oddDataSet1.IT_7A_G1(a).Type = "P" Then

                DB_oddDataSet1.IT_7A_G1(a).No_per_week = 3
                DB_oddDataSet1.IT_7A_G2(a).No_per_week = 3
                DB_oddDataSet1.IT_7A_G3(a).No_per_week = 3
                DB_oddDataSet1.IT_7B_G1(a).No_per_week = 3
                DB_oddDataSet1.IT_7B_G2(a).No_per_week = 3
                DB_oddDataSet1.IT_7B_G3(a).No_per_week = 3

            ElseIf DB_oddDataSet1.IT_7B_G1(a).Type = "T" Then

                DB_oddDataSet1.IT_7A_G1(a).No_per_week = 2
                DB_oddDataSet1.IT_7A_G2(a).No_per_week = 2
                DB_oddDataSet1.IT_7A_G3(a).No_per_week = 2
                DB_oddDataSet1.IT_7B_G1(a).No_per_week = 2
                DB_oddDataSet1.IT_7B_G2(a).No_per_week = 2
                DB_oddDataSet1.IT_7B_G3(a).No_per_week = 2

            End If

            DB_oddDataSet1.IT_7A_G1(a).Taken = 0
            DB_oddDataSet1.IT_7A_G2(a).Taken = 0
            DB_oddDataSet1.IT_7A_G3(a).Taken = 0
            DB_oddDataSet1.IT_7B_G1(a).Taken = 0
            DB_oddDataSet1.IT_7B_G2(a).Taken = 0
            DB_oddDataSet1.IT_7B_G3(a).Taken = 0

        Next a

        ' Semester 8

        For a = 0 To DB_evenDataSet1.IT_8A_G1.Count - 1

            If DB_evenDataSet1.IT_8A_G1(a).Type = "L" Or DB_evenDataSet1.IT_8A_G1(a).Type = "P" Then

                DB_evenDataSet1.IT_8A_G1(a).No_per_week = 3
                DB_evenDataSet1.IT_8A_G2(a).No_per_week = 3
                DB_evenDataSet1.IT_8A_G3(a).No_per_week = 3
                DB_evenDataSet1.IT_8B_G1(a).No_per_week = 3
                DB_evenDataSet1.IT_8B_G2(a).No_per_week = 3
                DB_evenDataSet1.IT_8B_G3(a).No_per_week = 3

            ElseIf DB_evenDataSet1.IT_8B_G1(a).Type = "T" Then

                DB_evenDataSet1.IT_8A_G1(a).No_per_week = 2
                DB_evenDataSet1.IT_8A_G2(a).No_per_week = 2
                DB_evenDataSet1.IT_8A_G3(a).No_per_week = 2
                DB_evenDataSet1.IT_8B_G1(a).No_per_week = 2
                DB_evenDataSet1.IT_8B_G2(a).No_per_week = 2
                DB_evenDataSet1.IT_8B_G3(a).No_per_week = 2

            End If

            DB_evenDataSet1.IT_8A_G1(a).Taken = 0
            DB_evenDataSet1.IT_8A_G2(a).Taken = 0
            DB_evenDataSet1.IT_8A_G3(a).Taken = 0
            DB_evenDataSet1.IT_8B_G1(a).Taken = 0
            DB_evenDataSet1.IT_8B_G2(a).Taken = 0
            DB_evenDataSet1.IT_8B_G3(a).Taken = 0

        Next a

        For a = 0 To DB_teacher_EVENDataSet1.Teachers.Count - 1

            ' Reset Even

            DB_teacher_EVENDataSet1.Teachers(a).Max_Hours = 16
            DB_teacher_EVENDataSet1.Teachers(a).Allotted = 0

            ' Reset Monday

            DB_teacher_EVENDataSet1.Teachers(a)._M_8_9 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_9_10 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_10_11 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_11_12 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_12_1 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_1_2 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_2_3 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_3_4 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_4_5 = False
            DB_teacher_EVENDataSet1.Teachers(a)._M_5_6 = False


            ' Reset Tuesday

            DB_teacher_EVENDataSet1.Teachers(a)._TU_8_9 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_9_10 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_10_11 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_11_12 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_12_1 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_1_2 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_2_3 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_3_4 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_4_5 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TU_5_6 = False

            ' Reset Wednessday

            DB_teacher_EVENDataSet1.Teachers(a)._W_8_9 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_9_10 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_10_11 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_11_12 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_12_1 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_1_2 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_2_3 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_3_4 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_4_5 = False
            DB_teacher_EVENDataSet1.Teachers(a)._W_5_6 = False

            ' Reset Thursday

            DB_teacher_EVENDataSet1.Teachers(a)._TH_8_9 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_9_10 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_10_11 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_11_12 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_12_1 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_1_2 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_2_3 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_3_4 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_4_5 = False
            DB_teacher_EVENDataSet1.Teachers(a)._TH_5_6 = False

            ' Reset Friday

            DB_teacher_EVENDataSet1.Teachers(a)._F_8_9 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_9_10 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_10_11 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_11_12 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_12_1 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_1_2 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_2_3 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_3_4 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_4_5 = False
            DB_teacher_EVENDataSet1.Teachers(a)._F_5_6 = False

        Next a

        For a = 0 To DB_teacher_ODDDataSet1.Teachers.Count - 1

            ' Reset Odd

            DB_teacher_ODDDataSet1.Teachers(a).Max_Hours = 16
            DB_teacher_ODDDataSet1.Teachers(a).Allotted = 0

            ' Reset Monday

            DB_teacher_ODDDataSet1.Teachers(a)._M_8_9 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_9_10 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_10_11 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_11_12 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_12_1 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_1_2 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_2_3 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_3_4 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_4_5 = False
            DB_teacher_ODDDataSet1.Teachers(a)._M_5_6 = False

            ' Reset Tuesday

            DB_teacher_ODDDataSet1.Teachers(a)._TU_8_9 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_9_10 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_10_11 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_11_12 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_12_1 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_1_2 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_2_3 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_3_4 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_4_5 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TU_5_6 = False

            ' Reset Wednessday

            DB_teacher_ODDDataSet1.Teachers(a)._W_8_9 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_9_10 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_10_11 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_11_12 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_12_1 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_1_2 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_2_3 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_3_4 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_4_5 = False
            DB_teacher_ODDDataSet1.Teachers(a)._W_5_6 = False

            ' Reset Thursday

            DB_teacher_ODDDataSet1.Teachers(a)._TH_8_9 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_9_10 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_10_11 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_11_12 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_12_1 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_1_2 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_2_3 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_3_4 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_4_5 = False
            DB_teacher_ODDDataSet1.Teachers(a)._TH_5_6 = False

            ' Reset Friday

            DB_teacher_ODDDataSet1.Teachers(a)._F_8_9 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_9_10 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_10_11 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_11_12 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_12_1 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_1_2 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_2_3 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_3_4 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_4_5 = False
            DB_teacher_ODDDataSet1.Teachers(a)._F_5_6 = False

        Next a
        For a = 0 To DB_room_EVENDataSet1.Rooms.Count - 1
            DB_room_EVENDataSet1.Rooms(a)._M_8_9 = False
            DB_room_EVENDataSet1.Rooms(a)._M_9_10 = False
            DB_room_EVENDataSet1.Rooms(a)._M_10_11 = False
            DB_room_EVENDataSet1.Rooms(a)._M_11_12 = False
            DB_room_EVENDataSet1.Rooms(a)._M_12_1 = False
            DB_room_EVENDataSet1.Rooms(a)._M_1_2 = False
            DB_room_EVENDataSet1.Rooms(a)._M_2_3 = False
            DB_room_EVENDataSet1.Rooms(a)._M_3_4 = False
            DB_room_EVENDataSet1.Rooms(a)._M_4_5 = False
            DB_room_EVENDataSet1.Rooms(a)._M_5_6 = False

            DB_room_EVENDataSet1.Rooms(a)._TU_8_9 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_9_10 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_10_11 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_11_12 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_12_1 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_1_2 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_2_3 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_3_4 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_4_5 = False
            DB_room_EVENDataSet1.Rooms(a)._TU_5_6 = False

            DB_room_EVENDataSet1.Rooms(a)._W_8_9 = False
            DB_room_EVENDataSet1.Rooms(a)._W_9_10 = False
            DB_room_EVENDataSet1.Rooms(a)._W_10_11 = False
            DB_room_EVENDataSet1.Rooms(a)._W_11_12 = False
            DB_room_EVENDataSet1.Rooms(a)._W_12_1 = False
            DB_room_EVENDataSet1.Rooms(a)._W_1_2 = False
            DB_room_EVENDataSet1.Rooms(a)._W_2_3 = False
            DB_room_EVENDataSet1.Rooms(a)._W_3_4 = False
            DB_room_EVENDataSet1.Rooms(a)._W_4_5 = False
            DB_room_EVENDataSet1.Rooms(a)._W_5_6 = False

            DB_room_EVENDataSet1.Rooms(a)._TH_8_9 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_9_10 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_10_11 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_11_12 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_12_1 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_1_2 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_2_3 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_3_4 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_4_5 = False
            DB_room_EVENDataSet1.Rooms(a)._TH_5_6 = False

            DB_room_EVENDataSet1.Rooms(a)._F_8_9 = False
            DB_room_EVENDataSet1.Rooms(a)._F_9_10 = False
            DB_room_EVENDataSet1.Rooms(a)._F_10_11 = False
            DB_room_EVENDataSet1.Rooms(a)._F_11_12 = False
            DB_room_EVENDataSet1.Rooms(a)._F_12_1 = False
            DB_room_EVENDataSet1.Rooms(a)._F_1_2 = False
            DB_room_EVENDataSet1.Rooms(a)._F_2_3 = False
            DB_room_EVENDataSet1.Rooms(a)._F_3_4 = False
            DB_room_EVENDataSet1.Rooms(a)._F_4_5 = False
            DB_room_EVENDataSet1.Rooms(a)._F_5_6 = False

        Next a
        For a = 0 To DB_room_ODDDataSet1.Rooms.Count - 1
            DB_room_ODDDataSet1.Rooms(a)._M_8_9 = False
            DB_room_ODDDataSet1.Rooms(a)._M_9_10 = False
            DB_room_ODDDataSet1.Rooms(a)._M_10_11 = False
            DB_room_ODDDataSet1.Rooms(a)._M_11_12 = False
            DB_room_ODDDataSet1.Rooms(a)._M_12_1 = False
            DB_room_ODDDataSet1.Rooms(a)._M_1_2 = False
            DB_room_ODDDataSet1.Rooms(a)._M_2_3 = False
            DB_room_ODDDataSet1.Rooms(a)._M_3_4 = False
            DB_room_ODDDataSet1.Rooms(a)._M_4_5 = False
            DB_room_ODDDataSet1.Rooms(a)._M_5_6 = False

            DB_room_ODDDataSet1.Rooms(a)._TU_8_9 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_9_10 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_10_11 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_11_12 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_12_1 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_1_2 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_2_3 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_3_4 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_4_5 = False
            DB_room_ODDDataSet1.Rooms(a)._TU_5_6 = False

            DB_room_ODDDataSet1.Rooms(a)._W_8_9 = False
            DB_room_ODDDataSet1.Rooms(a)._W_9_10 = False
            DB_room_ODDDataSet1.Rooms(a)._W_10_11 = False
            DB_room_ODDDataSet1.Rooms(a)._W_11_12 = False
            DB_room_ODDDataSet1.Rooms(a)._W_12_1 = False
            DB_room_ODDDataSet1.Rooms(a)._W_1_2 = False
            DB_room_ODDDataSet1.Rooms(a)._W_2_3 = False
            DB_room_ODDDataSet1.Rooms(a)._W_3_4 = False
            DB_room_ODDDataSet1.Rooms(a)._W_4_5 = False
            DB_room_ODDDataSet1.Rooms(a)._W_5_6 = False

            DB_room_ODDDataSet1.Rooms(a)._TH_8_9 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_9_10 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_10_11 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_11_12 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_12_1 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_1_2 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_2_3 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_3_4 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_4_5 = False
            DB_room_ODDDataSet1.Rooms(a)._TH_5_6 = False

            DB_room_ODDDataSet1.Rooms(a)._F_8_9 = False
            DB_room_ODDDataSet1.Rooms(a)._F_9_10 = False
            DB_room_ODDDataSet1.Rooms(a)._F_10_11 = False
            DB_room_ODDDataSet1.Rooms(a)._F_11_12 = False
            DB_room_ODDDataSet1.Rooms(a)._F_12_1 = False
            DB_room_ODDDataSet1.Rooms(a)._F_1_2 = False
            DB_room_ODDDataSet1.Rooms(a)._F_2_3 = False
            DB_room_ODDDataSet1.Rooms(a)._F_3_4 = False
            DB_room_ODDDataSet1.Rooms(a)._F_4_5 = False
            DB_room_ODDDataSet1.Rooms(a)._F_5_6 = False

        Next a
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If cmb101.Text <> "Odd or Even" And cmb102.Text <> "Semester" And cmb103.Text <> "Branch" And cmb104.Text <> "Section" And cmb105.Text <> "Group" Then
            If cmb102.Text = "3" Then
                If cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_3A_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_3A_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_3A_G3.Show()
                    End If
                ElseIf cmb104.Text = "Sec - B" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_3B_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_3B_g2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_3B_g3.Show()
                    End If
                End If
            ElseIf cmb102.Text = "4" Then
                If cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_4A_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_4A_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_4A_G3.Show()
                    End If
                ElseIf cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_4B_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_4B_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_4B_G3.Show()
                    End If
                End If
            ElseIf cmb102.Text = "5" Then
                If cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_5A_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_5A_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_5A_G3.Show()
                    End If
                ElseIf cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_5B_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_5B_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_5B_g3.Show()
                    End If
                End If
            ElseIf cmb102.Text = "6" Then
                If cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_6A_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_6A_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_6A_G3.Show()
                    End If
                ElseIf cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_6B_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_6B_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_6B_G3.Show()
                    End If
                End If
            ElseIf cmb102.Text = "7" Then
                If cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_7A_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_7A_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_7A_G3.Show()
                    End If
                ElseIf cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_7B_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_7B_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_7B_G3.Show()
                    End If
                End If
            ElseIf cmb102.Text = "8" Then
                If cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_8A_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_8A_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_8A_G3.Show()
                    End If
                ElseIf cmb104.Text = "Sec - A" Then
                    If cmb105.Text = "Group - 1" Then
                        frmIT_8B_G1.Show()
                    ElseIf cmb105.Text = "Group - 2" Then
                        frmIT_8B_G2.Show()
                    ElseIf cmb105.Text = "Group - 3" Then
                        frmIT_8B_G3.Show()
                    End If
                End If
                Me.Hide()
            End If
        Else
            MsgBox("Error !!! Please fill in all prerequisite feilds...!!!", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub cmb101_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb101.SelectedIndexChanged
        If cmb101.Text = "Odd" Then
            cmb102.Items.Clear()
            '     cmb102.Items.Add(1)
            cmb102.Items.Add(3)
            cmb102.Items.Add(5)
            cmb102.Items.Add(7)
        Else
            cmb102.Items.Clear()
            '      cmb102.Items.Add(2)
            cmb102.Items.Add(4)
            cmb102.Items.Add(6)
            cmb102.Items.Add(8)
        End If
    End Sub

    Public Sub ExcelExport() Handles ExcelToolStripMenuItem.Click

        ' Declare Variables
        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        Dim oRng As Excel.Range

        ' Start Excel and get Application object.
        oXL = CreateObject("Excel.Application")
        oXL.Visible = True


        If cmb101.Text = "ODD" Then

            Dim Sheet_IT_3A_G1 As Excel.Worksheet
            Dim Sheet_IT_3A_G2 As Excel.Worksheet
            Dim Sheet_IT_3A_G3 As Excel.Worksheet
            Dim Sheet_IT_3B_G1 As Excel.Worksheet
            Dim Sheet_IT_3B_G2 As Excel.Worksheet
            Dim Sheet_IT_3B_G3 As Excel.Worksheet

            Dim Sheet_IT_5A_G1 As Excel.Worksheet
            Dim Sheet_IT_5A_G2 As Excel.Worksheet
            Dim Sheet_IT_5A_G3 As Excel.Worksheet
            Dim Sheet_IT_5B_G1 As Excel.Worksheet
            Dim Sheet_IT_5B_G2 As Excel.Worksheet
            Dim Sheet_IT_5B_G3 As Excel.Worksheet

            Dim Sheet_IT_7A_G1 As Excel.Worksheet
            Dim Sheet_IT_7A_G2 As Excel.Worksheet
            Dim Sheet_IT_7A_G3 As Excel.Worksheet
            Dim Sheet_IT_7B_G1 As Excel.Worksheet
            Dim Sheet_IT_7B_G2 As Excel.Worksheet
            Dim Sheet_IT_7B_G3 As Excel.Worksheet

            ' Get new workbooks.

            oWB = oXL.Workbooks.Add
            Sheet_IT_3A_G1 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_3A_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_3A_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_3A_G1.Cells(1, 1).Value = ""
            Sheet_IT_3A_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_3A_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_3A_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_3A_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_3A_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_3A_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_3A_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_3A_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_3A_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_3A_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_3A_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_3A_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_3A_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_3A_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_3A_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_3A_G1.Cells(2, 2).Value = frmIT_3A_G1.cmb1.Text
            Sheet_IT_3A_G1.Cells(3, 2).Value = frmIT_3A_G1.cmb2.Text
            Sheet_IT_3A_G1.Cells(4, 2).Value = frmIT_3A_G1.cmb3.Text
            Sheet_IT_3A_G1.Cells(5, 2).Value = frmIT_3A_G1.cmb4.Text
            Sheet_IT_3A_G1.Cells(6, 2).Value = frmIT_3A_G1.cmb5.Text
            Sheet_IT_3A_G1.Cells(7, 2).Value = frmIT_3A_G1.cmb6.Text
            Sheet_IT_3A_G1.Cells(8, 2).Value = frmIT_3A_G1.cmb7.Text
            Sheet_IT_3A_G1.Cells(9, 2).Value = frmIT_3A_G1.cmb8.Text
            Sheet_IT_3A_G1.Cells(10, 2).Value = frmIT_3A_G1.cmb9.Text
            Sheet_IT_3A_G1.Cells(11, 2).Value = frmIT_3A_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_3A_G1.Cells(2, 3).Value = frmIT_3A_G1.cmb21.Text
            Sheet_IT_3A_G1.Cells(3, 3).Value = frmIT_3A_G1.cmb22.Text
            Sheet_IT_3A_G1.Cells(4, 3).Value = frmIT_3A_G1.cmb23.Text
            Sheet_IT_3A_G1.Cells(5, 3).Value = frmIT_3A_G1.cmb24.Text
            Sheet_IT_3A_G1.Cells(6, 3).Value = frmIT_3A_G1.cmb25.Text
            Sheet_IT_3A_G1.Cells(7, 3).Value = frmIT_3A_G1.cmb26.Text
            Sheet_IT_3A_G1.Cells(8, 3).Value = frmIT_3A_G1.cmb27.Text
            Sheet_IT_3A_G1.Cells(9, 3).Value = frmIT_3A_G1.cmb28.Text
            Sheet_IT_3A_G1.Cells(10, 3).Value = frmIT_3A_G1.cmb29.Text
            Sheet_IT_3A_G1.Cells(11, 3).Value = frmIT_3A_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_3A_G1.Cells(2, 4).Value = frmIT_3A_G1.cmb41.Text
            Sheet_IT_3A_G1.Cells(3, 4).Value = frmIT_3A_G1.cmb42.Text
            Sheet_IT_3A_G1.Cells(4, 4).Value = frmIT_3A_G1.cmb43.Text
            Sheet_IT_3A_G1.Cells(5, 4).Value = frmIT_3A_G1.cmb44.Text
            Sheet_IT_3A_G1.Cells(6, 4).Value = frmIT_3A_G1.cmb45.Text
            Sheet_IT_3A_G1.Cells(7, 4).Value = frmIT_3A_G1.cmb46.Text
            Sheet_IT_3A_G1.Cells(8, 4).Value = frmIT_3A_G1.cmb47.Text
            Sheet_IT_3A_G1.Cells(9, 4).Value = frmIT_3A_G1.cmb48.Text
            Sheet_IT_3A_G1.Cells(10, 4).Value = frmIT_3A_G1.cmb49.Text
            Sheet_IT_3A_G1.Cells(11, 4).Value = frmIT_3A_G1.cmb50.Text

            ' Thursday

            Sheet_IT_3A_G1.Cells(2, 5).Value = frmIT_3A_G1.cmb61.Text
            Sheet_IT_3A_G1.Cells(3, 5).Value = frmIT_3A_G1.cmb62.Text
            Sheet_IT_3A_G1.Cells(4, 5).Value = frmIT_3A_G1.cmb63.Text
            Sheet_IT_3A_G1.Cells(5, 5).Value = frmIT_3A_G1.cmb64.Text
            Sheet_IT_3A_G1.Cells(6, 5).Value = frmIT_3A_G1.cmb65.Text
            Sheet_IT_3A_G1.Cells(7, 5).Value = frmIT_3A_G1.cmb66.Text
            Sheet_IT_3A_G1.Cells(8, 5).Value = frmIT_3A_G1.cmb67.Text
            Sheet_IT_3A_G1.Cells(9, 5).Value = frmIT_3A_G1.cmb68.Text
            Sheet_IT_3A_G1.Cells(10, 5).Value = frmIT_3A_G1.cmb69.Text
            Sheet_IT_3A_G1.Cells(11, 5).Value = frmIT_3A_G1.cmb70.Text

            ' Friday

            Sheet_IT_3A_G1.Cells(2, 6).Value = frmIT_3A_G1.cmb81.Text
            Sheet_IT_3A_G1.Cells(3, 6).Value = frmIT_3A_G1.cmb82.Text
            Sheet_IT_3A_G1.Cells(4, 6).Value = frmIT_3A_G1.cmb83.Text
            Sheet_IT_3A_G1.Cells(5, 6).Value = frmIT_3A_G1.cmb84.Text
            Sheet_IT_3A_G1.Cells(6, 6).Value = frmIT_3A_G1.cmb85.Text
            Sheet_IT_3A_G1.Cells(7, 6).Value = frmIT_3A_G1.cmb86.Text
            Sheet_IT_3A_G1.Cells(8, 6).Value = frmIT_3A_G1.cmb87.Text
            Sheet_IT_3A_G1.Cells(9, 6).Value = frmIT_3A_G1.cmb88.Text
            Sheet_IT_3A_G1.Cells(10, 6).Value = frmIT_3A_G1.cmb89.Text
            Sheet_IT_3A_G1.Cells(11, 6).Value = frmIT_3A_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_3A_G2 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_3A_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_3A_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_3A_G2.Cells(1, 1).Value = ""
            Sheet_IT_3A_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_3A_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_3A_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_3A_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_3A_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_3A_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_3A_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_3A_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_3A_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_3A_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_3A_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_3A_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_3A_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_3A_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_3A_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_3A_G2.Cells(2, 2).Value = frmIT_3A_G2.cmb1.Text
            Sheet_IT_3A_G2.Cells(3, 2).Value = frmIT_3A_G2.cmb2.Text
            Sheet_IT_3A_G2.Cells(4, 2).Value = frmIT_3A_G2.cmb3.Text
            Sheet_IT_3A_G2.Cells(5, 2).Value = frmIT_3A_G2.cmb4.Text
            Sheet_IT_3A_G2.Cells(6, 2).Value = frmIT_3A_G2.cmb5.Text
            Sheet_IT_3A_G2.Cells(7, 2).Value = frmIT_3A_G2.cmb6.Text
            Sheet_IT_3A_G2.Cells(8, 2).Value = frmIT_3A_G2.cmb7.Text
            Sheet_IT_3A_G2.Cells(9, 2).Value = frmIT_3A_G2.cmb8.Text
            Sheet_IT_3A_G2.Cells(10, 2).Value = frmIT_3A_G2.cmb9.Text
            Sheet_IT_3A_G2.Cells(11, 2).Value = frmIT_3A_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_3A_G2.Cells(2, 3).Value = frmIT_3A_G2.cmb21.Text
            Sheet_IT_3A_G2.Cells(3, 3).Value = frmIT_3A_G2.cmb22.Text
            Sheet_IT_3A_G2.Cells(4, 3).Value = frmIT_3A_G2.cmb23.Text
            Sheet_IT_3A_G2.Cells(5, 3).Value = frmIT_3A_G2.cmb24.Text
            Sheet_IT_3A_G2.Cells(6, 3).Value = frmIT_3A_G2.cmb25.Text
            Sheet_IT_3A_G2.Cells(7, 3).Value = frmIT_3A_G2.cmb26.Text
            Sheet_IT_3A_G2.Cells(8, 3).Value = frmIT_3A_G2.cmb27.Text
            Sheet_IT_3A_G2.Cells(9, 3).Value = frmIT_3A_G2.cmb28.Text
            Sheet_IT_3A_G2.Cells(10, 3).Value = frmIT_3A_G2.cmb29.Text
            Sheet_IT_3A_G2.Cells(11, 3).Value = frmIT_3A_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_3A_G2.Cells(2, 4).Value = frmIT_3A_G2.cmb41.Text
            Sheet_IT_3A_G2.Cells(3, 4).Value = frmIT_3A_G2.cmb42.Text
            Sheet_IT_3A_G2.Cells(4, 4).Value = frmIT_3A_G2.cmb43.Text
            Sheet_IT_3A_G2.Cells(5, 4).Value = frmIT_3A_G2.cmb44.Text
            Sheet_IT_3A_G2.Cells(6, 4).Value = frmIT_3A_G2.cmb45.Text
            Sheet_IT_3A_G2.Cells(7, 4).Value = frmIT_3A_G2.cmb46.Text
            Sheet_IT_3A_G2.Cells(8, 4).Value = frmIT_3A_G2.cmb47.Text
            Sheet_IT_3A_G2.Cells(9, 4).Value = frmIT_3A_G2.cmb48.Text
            Sheet_IT_3A_G2.Cells(10, 4).Value = frmIT_3A_G2.cmb49.Text
            Sheet_IT_3A_G2.Cells(11, 4).Value = frmIT_3A_G2.cmb50.Text

            ' Thursday

            Sheet_IT_3A_G2.Cells(2, 5).Value = frmIT_3A_G2.cmb61.Text
            Sheet_IT_3A_G2.Cells(3, 5).Value = frmIT_3A_G2.cmb62.Text
            Sheet_IT_3A_G2.Cells(4, 5).Value = frmIT_3A_G2.cmb63.Text
            Sheet_IT_3A_G2.Cells(5, 5).Value = frmIT_3A_G2.cmb64.Text
            Sheet_IT_3A_G2.Cells(6, 5).Value = frmIT_3A_G2.cmb65.Text
            Sheet_IT_3A_G2.Cells(7, 5).Value = frmIT_3A_G2.cmb66.Text
            Sheet_IT_3A_G2.Cells(8, 5).Value = frmIT_3A_G2.cmb67.Text
            Sheet_IT_3A_G2.Cells(9, 5).Value = frmIT_3A_G2.cmb68.Text
            Sheet_IT_3A_G2.Cells(10, 5).Value = frmIT_3A_G2.cmb69.Text
            Sheet_IT_3A_G2.Cells(11, 5).Value = frmIT_3A_G2.cmb70.Text

            ' Friday

            Sheet_IT_3A_G2.Cells(2, 6).Value = frmIT_3A_G2.cmb81.Text
            Sheet_IT_3A_G2.Cells(3, 6).Value = frmIT_3A_G2.cmb82.Text
            Sheet_IT_3A_G2.Cells(4, 6).Value = frmIT_3A_G2.cmb83.Text
            Sheet_IT_3A_G2.Cells(5, 6).Value = frmIT_3A_G2.cmb84.Text
            Sheet_IT_3A_G2.Cells(6, 6).Value = frmIT_3A_G2.cmb85.Text
            Sheet_IT_3A_G2.Cells(7, 6).Value = frmIT_3A_G2.cmb86.Text
            Sheet_IT_3A_G2.Cells(8, 6).Value = frmIT_3A_G2.cmb87.Text
            Sheet_IT_3A_G2.Cells(9, 6).Value = frmIT_3A_G2.cmb88.Text
            Sheet_IT_3A_G2.Cells(10, 6).Value = frmIT_3A_G2.cmb89.Text
            Sheet_IT_3A_G2.Cells(11, 6).Value = frmIT_3A_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_3A_G3 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_3A_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_3A_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_3A_G3.Cells(1, 1).Value = ""
            Sheet_IT_3A_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_3A_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_3A_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_3A_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_3A_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_3A_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_3A_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_3A_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_3A_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_3A_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_3A_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_3A_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_3A_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_3A_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_3A_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_3A_G3.Cells(2, 2).Value = frmIT_3A_G3.cmb1.Text
            Sheet_IT_3A_G3.Cells(3, 2).Value = frmIT_3A_G3.cmb2.Text
            Sheet_IT_3A_G3.Cells(4, 2).Value = frmIT_3A_G3.cmb3.Text
            Sheet_IT_3A_G3.Cells(5, 2).Value = frmIT_3A_G3.cmb4.Text
            Sheet_IT_3A_G3.Cells(6, 2).Value = frmIT_3A_G3.cmb5.Text
            Sheet_IT_3A_G3.Cells(7, 2).Value = frmIT_3A_G3.cmb6.Text
            Sheet_IT_3A_G3.Cells(8, 2).Value = frmIT_3A_G3.cmb7.Text
            Sheet_IT_3A_G3.Cells(9, 2).Value = frmIT_3A_G3.cmb8.Text
            Sheet_IT_3A_G3.Cells(10, 2).Value = frmIT_3A_G3.cmb9.Text
            Sheet_IT_3A_G3.Cells(11, 2).Value = frmIT_3A_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_3A_G3.Cells(2, 3).Value = frmIT_3A_G3.cmb21.Text
            Sheet_IT_3A_G3.Cells(3, 3).Value = frmIT_3A_G3.cmb22.Text
            Sheet_IT_3A_G3.Cells(4, 3).Value = frmIT_3A_G3.cmb23.Text
            Sheet_IT_3A_G3.Cells(5, 3).Value = frmIT_3A_G3.cmb24.Text
            Sheet_IT_3A_G3.Cells(6, 3).Value = frmIT_3A_G3.cmb25.Text
            Sheet_IT_3A_G3.Cells(7, 3).Value = frmIT_3A_G3.cmb26.Text
            Sheet_IT_3A_G3.Cells(8, 3).Value = frmIT_3A_G3.cmb27.Text
            Sheet_IT_3A_G3.Cells(9, 3).Value = frmIT_3A_G3.cmb28.Text
            Sheet_IT_3A_G3.Cells(10, 3).Value = frmIT_3A_G3.cmb29.Text
            Sheet_IT_3A_G3.Cells(11, 3).Value = frmIT_3A_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_3A_G3.Cells(2, 4).Value = frmIT_3A_G3.cmb41.Text
            Sheet_IT_3A_G3.Cells(3, 4).Value = frmIT_3A_G3.cmb42.Text
            Sheet_IT_3A_G3.Cells(4, 4).Value = frmIT_3A_G3.cmb43.Text
            Sheet_IT_3A_G3.Cells(5, 4).Value = frmIT_3A_G3.cmb44.Text
            Sheet_IT_3A_G3.Cells(6, 4).Value = frmIT_3A_G3.cmb45.Text
            Sheet_IT_3A_G3.Cells(7, 4).Value = frmIT_3A_G3.cmb46.Text
            Sheet_IT_3A_G3.Cells(8, 4).Value = frmIT_3A_G3.cmb47.Text
            Sheet_IT_3A_G3.Cells(9, 4).Value = frmIT_3A_G3.cmb48.Text
            Sheet_IT_3A_G3.Cells(10, 4).Value = frmIT_3A_G3.cmb49.Text
            Sheet_IT_3A_G3.Cells(11, 4).Value = frmIT_3A_G3.cmb50.Text

            ' Thursday

            Sheet_IT_3A_G3.Cells(2, 5).Value = frmIT_3A_G3.cmb61.Text
            Sheet_IT_3A_G3.Cells(3, 5).Value = frmIT_3A_G3.cmb62.Text
            Sheet_IT_3A_G3.Cells(4, 5).Value = frmIT_3A_G3.cmb63.Text
            Sheet_IT_3A_G3.Cells(5, 5).Value = frmIT_3A_G3.cmb64.Text
            Sheet_IT_3A_G3.Cells(6, 5).Value = frmIT_3A_G3.cmb65.Text
            Sheet_IT_3A_G3.Cells(7, 5).Value = frmIT_3A_G3.cmb66.Text
            Sheet_IT_3A_G3.Cells(8, 5).Value = frmIT_3A_G3.cmb67.Text
            Sheet_IT_3A_G3.Cells(9, 5).Value = frmIT_3A_G3.cmb68.Text
            Sheet_IT_3A_G3.Cells(10, 5).Value = frmIT_3A_G3.cmb69.Text
            Sheet_IT_3A_G3.Cells(11, 5).Value = frmIT_3A_G3.cmb70.Text

            ' Friday

            Sheet_IT_3A_G3.Cells(2, 6).Value = frmIT_3A_G3.cmb81.Text
            Sheet_IT_3A_G3.Cells(3, 6).Value = frmIT_3A_G3.cmb82.Text
            Sheet_IT_3A_G3.Cells(4, 6).Value = frmIT_3A_G3.cmb83.Text
            Sheet_IT_3A_G3.Cells(5, 6).Value = frmIT_3A_G3.cmb84.Text
            Sheet_IT_3A_G3.Cells(6, 6).Value = frmIT_3A_G3.cmb85.Text
            Sheet_IT_3A_G3.Cells(7, 6).Value = frmIT_3A_G3.cmb86.Text
            Sheet_IT_3A_G3.Cells(8, 6).Value = frmIT_3A_G3.cmb87.Text
            Sheet_IT_3A_G3.Cells(9, 6).Value = frmIT_3A_G3.cmb88.Text
            Sheet_IT_3A_G3.Cells(10, 6).Value = frmIT_3A_G3.cmb89.Text
            Sheet_IT_3A_G3.Cells(11, 6).Value = frmIT_3A_G3.cmb90.Text
            ''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_3B_G1 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_3B_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_3B_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_3B_G1.Cells(1, 1).Value = ""
            Sheet_IT_3B_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_3B_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_3B_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_3B_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_3B_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_3B_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_3B_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_3B_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_3B_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_3B_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_3B_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_3B_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_3B_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_3B_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_3B_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_3B_G1.Cells(2, 2).Value = frmIT_3B_G1.cmb1.Text
            Sheet_IT_3B_G1.Cells(3, 2).Value = frmIT_3B_G1.cmb2.Text
            Sheet_IT_3B_G1.Cells(4, 2).Value = frmIT_3B_G1.cmb3.Text
            Sheet_IT_3B_G1.Cells(5, 2).Value = frmIT_3B_G1.cmb4.Text
            Sheet_IT_3B_G1.Cells(6, 2).Value = frmIT_3B_G1.cmb5.Text
            Sheet_IT_3B_G1.Cells(7, 2).Value = frmIT_3B_G1.cmb6.Text
            Sheet_IT_3B_G1.Cells(8, 2).Value = frmIT_3B_G1.cmb7.Text
            Sheet_IT_3B_G1.Cells(9, 2).Value = frmIT_3B_G1.cmb8.Text
            Sheet_IT_3B_G1.Cells(10, 2).Value = frmIT_3B_G1.cmb9.Text
            Sheet_IT_3B_G1.Cells(11, 2).Value = frmIT_3B_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_3B_G1.Cells(2, 3).Value = frmIT_3B_G1.cmb21.Text
            Sheet_IT_3B_G1.Cells(3, 3).Value = frmIT_3B_G1.cmb22.Text
            Sheet_IT_3B_G1.Cells(4, 3).Value = frmIT_3B_G1.cmb23.Text
            Sheet_IT_3B_G1.Cells(5, 3).Value = frmIT_3B_G1.cmb24.Text
            Sheet_IT_3B_G1.Cells(6, 3).Value = frmIT_3B_G1.cmb25.Text
            Sheet_IT_3B_G1.Cells(7, 3).Value = frmIT_3B_G1.cmb26.Text
            Sheet_IT_3B_G1.Cells(8, 3).Value = frmIT_3B_G1.cmb27.Text
            Sheet_IT_3B_G1.Cells(9, 3).Value = frmIT_3B_G1.cmb28.Text
            Sheet_IT_3B_G1.Cells(10, 3).Value = frmIT_3B_G1.cmb29.Text
            Sheet_IT_3B_G1.Cells(11, 3).Value = frmIT_3B_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_3B_G1.Cells(2, 4).Value = frmIT_3B_G1.cmb41.Text
            Sheet_IT_3B_G1.Cells(3, 4).Value = frmIT_3B_G1.cmb42.Text
            Sheet_IT_3B_G1.Cells(4, 4).Value = frmIT_3B_G1.cmb43.Text
            Sheet_IT_3B_G1.Cells(5, 4).Value = frmIT_3B_G1.cmb44.Text
            Sheet_IT_3B_G1.Cells(6, 4).Value = frmIT_3B_G1.cmb45.Text
            Sheet_IT_3B_G1.Cells(7, 4).Value = frmIT_3B_G1.cmb46.Text
            Sheet_IT_3B_G1.Cells(8, 4).Value = frmIT_3B_G1.cmb47.Text
            Sheet_IT_3B_G1.Cells(9, 4).Value = frmIT_3B_G1.cmb48.Text
            Sheet_IT_3B_G1.Cells(10, 4).Value = frmIT_3B_G1.cmb49.Text
            Sheet_IT_3B_G1.Cells(11, 4).Value = frmIT_3B_G1.cmb50.Text

            ' Thursday

            Sheet_IT_3B_G1.Cells(2, 5).Value = frmIT_3B_G1.cmb61.Text
            Sheet_IT_3B_G1.Cells(3, 5).Value = frmIT_3B_G1.cmb62.Text
            Sheet_IT_3B_G1.Cells(4, 5).Value = frmIT_3B_G1.cmb63.Text
            Sheet_IT_3B_G1.Cells(5, 5).Value = frmIT_3B_G1.cmb64.Text
            Sheet_IT_3B_G1.Cells(6, 5).Value = frmIT_3B_G1.cmb65.Text
            Sheet_IT_3B_G1.Cells(7, 5).Value = frmIT_3B_G1.cmb66.Text
            Sheet_IT_3B_G1.Cells(8, 5).Value = frmIT_3B_G1.cmb67.Text
            Sheet_IT_3B_G1.Cells(9, 5).Value = frmIT_3B_G1.cmb68.Text
            Sheet_IT_3B_G1.Cells(10, 5).Value = frmIT_3B_G1.cmb69.Text
            Sheet_IT_3B_G1.Cells(11, 5).Value = frmIT_3B_G1.cmb70.Text

            ' Friday

            Sheet_IT_3B_G1.Cells(2, 6).Value = frmIT_3B_G1.cmb81.Text
            Sheet_IT_3B_G1.Cells(3, 6).Value = frmIT_3B_G1.cmb82.Text
            Sheet_IT_3B_G1.Cells(4, 6).Value = frmIT_3B_G1.cmb83.Text
            Sheet_IT_3B_G1.Cells(5, 6).Value = frmIT_3B_G1.cmb84.Text
            Sheet_IT_3B_G1.Cells(6, 6).Value = frmIT_3B_G1.cmb85.Text
            Sheet_IT_3B_G1.Cells(7, 6).Value = frmIT_3B_G1.cmb86.Text
            Sheet_IT_3B_G1.Cells(8, 6).Value = frmIT_3B_G1.cmb87.Text
            Sheet_IT_3B_G1.Cells(9, 6).Value = frmIT_3B_G1.cmb88.Text
            Sheet_IT_3B_G1.Cells(10, 6).Value = frmIT_3B_G1.cmb89.Text
            Sheet_IT_3B_G1.Cells(11, 6).Value = frmIT_3B_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_3B_G2 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_3B_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_3B_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_3B_G2.Cells(1, 1).Value = ""
            Sheet_IT_3B_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_3B_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_3B_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_3B_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_3B_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_3B_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_3B_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_3B_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_3B_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_3B_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_3B_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_3B_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_3B_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_3B_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_3B_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_3B_G2.Cells(2, 2).Value = frmIT_3B_g2.cmb1.Text
            Sheet_IT_3B_G2.Cells(3, 2).Value = frmIT_3B_g2.cmb2.Text
            Sheet_IT_3B_G2.Cells(4, 2).Value = frmIT_3B_g2.cmb3.Text
            Sheet_IT_3B_G2.Cells(5, 2).Value = frmIT_3B_g2.cmb4.Text
            Sheet_IT_3B_G2.Cells(6, 2).Value = frmIT_3B_g2.cmb5.Text
            Sheet_IT_3B_G2.Cells(7, 2).Value = frmIT_3B_g2.cmb6.Text
            Sheet_IT_3B_G2.Cells(8, 2).Value = frmIT_3B_g2.cmb7.Text
            Sheet_IT_3B_G2.Cells(9, 2).Value = frmIT_3B_g2.cmb8.Text
            Sheet_IT_3B_G2.Cells(10, 2).Value = frmIT_3B_g2.cmb9.Text
            Sheet_IT_3B_G2.Cells(11, 2).Value = frmIT_3B_g2.cmb10.Text

            ' Tuesday

            Sheet_IT_3B_G2.Cells(2, 3).Value = frmIT_3B_g2.cmb21.Text
            Sheet_IT_3B_G2.Cells(3, 3).Value = frmIT_3B_g2.cmb22.Text
            Sheet_IT_3B_G2.Cells(4, 3).Value = frmIT_3B_g2.cmb23.Text
            Sheet_IT_3B_G2.Cells(5, 3).Value = frmIT_3B_g2.cmb24.Text
            Sheet_IT_3B_G2.Cells(6, 3).Value = frmIT_3B_g2.cmb25.Text
            Sheet_IT_3B_G2.Cells(7, 3).Value = frmIT_3B_g2.cmb26.Text
            Sheet_IT_3B_G2.Cells(8, 3).Value = frmIT_3B_g2.cmb27.Text
            Sheet_IT_3B_G2.Cells(9, 3).Value = frmIT_3B_g2.cmb28.Text
            Sheet_IT_3B_G2.Cells(10, 3).Value = frmIT_3B_g2.cmb29.Text
            Sheet_IT_3B_G2.Cells(11, 3).Value = frmIT_3B_g2.cmb30.Text

            ' Wednessday

            Sheet_IT_3B_G2.Cells(2, 4).Value = frmIT_3B_g2.cmb41.Text
            Sheet_IT_3B_G2.Cells(3, 4).Value = frmIT_3B_g2.cmb42.Text
            Sheet_IT_3B_G2.Cells(4, 4).Value = frmIT_3B_g2.cmb43.Text
            Sheet_IT_3B_G2.Cells(5, 4).Value = frmIT_3B_g2.cmb44.Text
            Sheet_IT_3B_G2.Cells(6, 4).Value = frmIT_3B_g2.cmb45.Text
            Sheet_IT_3B_G2.Cells(7, 4).Value = frmIT_3B_g2.cmb46.Text
            Sheet_IT_3B_G2.Cells(8, 4).Value = frmIT_3B_g2.cmb47.Text
            Sheet_IT_3B_G2.Cells(9, 4).Value = frmIT_3B_g2.cmb48.Text
            Sheet_IT_3B_G2.Cells(10, 4).Value = frmIT_3B_g2.cmb49.Text
            Sheet_IT_3B_G2.Cells(11, 4).Value = frmIT_3B_g2.cmb50.Text

            ' Thursday

            Sheet_IT_3B_G2.Cells(2, 5).Value = frmIT_3B_g2.cmb61.Text
            Sheet_IT_3B_G2.Cells(3, 5).Value = frmIT_3B_g2.cmb62.Text
            Sheet_IT_3B_G2.Cells(4, 5).Value = frmIT_3B_g2.cmb63.Text
            Sheet_IT_3B_G2.Cells(5, 5).Value = frmIT_3B_g2.cmb64.Text
            Sheet_IT_3B_G2.Cells(6, 5).Value = frmIT_3B_g2.cmb65.Text
            Sheet_IT_3B_G2.Cells(7, 5).Value = frmIT_3B_g2.cmb66.Text
            Sheet_IT_3B_G2.Cells(8, 5).Value = frmIT_3B_g2.cmb67.Text
            Sheet_IT_3B_G2.Cells(9, 5).Value = frmIT_3B_g2.cmb68.Text
            Sheet_IT_3B_G2.Cells(10, 5).Value = frmIT_3B_g2.cmb69.Text
            Sheet_IT_3B_G2.Cells(11, 5).Value = frmIT_3B_g2.cmb70.Text

            ' Friday

            Sheet_IT_3B_G2.Cells(2, 6).Value = frmIT_3B_g2.cmb81.Text
            Sheet_IT_3B_G2.Cells(3, 6).Value = frmIT_3B_g2.cmb82.Text
            Sheet_IT_3B_G2.Cells(4, 6).Value = frmIT_3B_g2.cmb83.Text
            Sheet_IT_3B_G2.Cells(5, 6).Value = frmIT_3B_g2.cmb84.Text
            Sheet_IT_3B_G2.Cells(6, 6).Value = frmIT_3B_g2.cmb85.Text
            Sheet_IT_3B_G2.Cells(7, 6).Value = frmIT_3B_g2.cmb86.Text
            Sheet_IT_3B_G2.Cells(8, 6).Value = frmIT_3B_g2.cmb87.Text
            Sheet_IT_3B_G2.Cells(9, 6).Value = frmIT_3B_g2.cmb88.Text
            Sheet_IT_3B_G2.Cells(10, 6).Value = frmIT_3B_g2.cmb89.Text
            Sheet_IT_3B_G2.Cells(11, 6).Value = frmIT_3B_g2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_3B_G3 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_3B_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_3B_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_3B_G3.Cells(1, 1).Value = ""
            Sheet_IT_3B_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_3B_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_3B_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_3B_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_3B_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_3B_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_3B_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_3B_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_3B_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_3B_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_3B_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_3B_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_3B_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_3B_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_3B_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_3B_G3.Cells(2, 2).Value = frmIT_3B_g3.cmb1.Text
            Sheet_IT_3B_G3.Cells(3, 2).Value = frmIT_3B_g3.cmb2.Text
            Sheet_IT_3B_G3.Cells(4, 2).Value = frmIT_3B_g3.cmb3.Text
            Sheet_IT_3B_G3.Cells(5, 2).Value = frmIT_3B_g3.cmb4.Text
            Sheet_IT_3B_G3.Cells(6, 2).Value = frmIT_3B_g3.cmb5.Text
            Sheet_IT_3B_G3.Cells(7, 2).Value = frmIT_3B_g3.cmb6.Text
            Sheet_IT_3B_G3.Cells(8, 2).Value = frmIT_3B_g3.cmb7.Text
            Sheet_IT_3B_G3.Cells(9, 2).Value = frmIT_3B_g3.cmb8.Text
            Sheet_IT_3B_G3.Cells(10, 2).Value = frmIT_3B_g3.cmb9.Text
            Sheet_IT_3B_G3.Cells(11, 2).Value = frmIT_3B_g3.cmb10.Text

            ' Tuesday

            Sheet_IT_3B_G3.Cells(2, 3).Value = frmIT_3B_g3.cmb21.Text
            Sheet_IT_3B_G3.Cells(3, 3).Value = frmIT_3B_g3.cmb22.Text
            Sheet_IT_3B_G3.Cells(4, 3).Value = frmIT_3B_g3.cmb23.Text
            Sheet_IT_3B_G3.Cells(5, 3).Value = frmIT_3B_g3.cmb24.Text
            Sheet_IT_3B_G3.Cells(6, 3).Value = frmIT_3B_g3.cmb25.Text
            Sheet_IT_3B_G3.Cells(7, 3).Value = frmIT_3B_g3.cmb26.Text
            Sheet_IT_3B_G3.Cells(8, 3).Value = frmIT_3B_g3.cmb27.Text
            Sheet_IT_3B_G3.Cells(9, 3).Value = frmIT_3B_g3.cmb28.Text
            Sheet_IT_3B_G3.Cells(10, 3).Value = frmIT_3B_g3.cmb29.Text
            Sheet_IT_3B_G3.Cells(11, 3).Value = frmIT_3B_g3.cmb30.Text

            ' Wednessday

            Sheet_IT_3B_G3.Cells(2, 4).Value = frmIT_3B_g3.cmb41.Text
            Sheet_IT_3B_G3.Cells(3, 4).Value = frmIT_3B_g3.cmb42.Text
            Sheet_IT_3B_G3.Cells(4, 4).Value = frmIT_3B_g3.cmb43.Text
            Sheet_IT_3B_G3.Cells(5, 4).Value = frmIT_3B_g3.cmb44.Text
            Sheet_IT_3B_G3.Cells(6, 4).Value = frmIT_3B_g3.cmb45.Text
            Sheet_IT_3B_G3.Cells(7, 4).Value = frmIT_3B_g3.cmb46.Text
            Sheet_IT_3B_G3.Cells(8, 4).Value = frmIT_3B_g3.cmb47.Text
            Sheet_IT_3B_G3.Cells(9, 4).Value = frmIT_3B_g3.cmb48.Text
            Sheet_IT_3B_G3.Cells(10, 4).Value = frmIT_3B_g3.cmb49.Text
            Sheet_IT_3B_G3.Cells(11, 4).Value = frmIT_3B_g3.cmb50.Text

            ' Thursday

            Sheet_IT_3B_G3.Cells(2, 5).Value = frmIT_3B_g3.cmb61.Text
            Sheet_IT_3B_G3.Cells(3, 5).Value = frmIT_3B_g3.cmb62.Text
            Sheet_IT_3B_G3.Cells(4, 5).Value = frmIT_3B_g3.cmb63.Text
            Sheet_IT_3B_G3.Cells(5, 5).Value = frmIT_3B_g3.cmb64.Text
            Sheet_IT_3B_G3.Cells(6, 5).Value = frmIT_3B_g3.cmb65.Text
            Sheet_IT_3B_G3.Cells(7, 5).Value = frmIT_3B_g3.cmb66.Text
            Sheet_IT_3B_G3.Cells(8, 5).Value = frmIT_3B_g3.cmb67.Text
            Sheet_IT_3B_G3.Cells(9, 5).Value = frmIT_3B_g3.cmb68.Text
            Sheet_IT_3B_G3.Cells(10, 5).Value = frmIT_3B_g3.cmb69.Text
            Sheet_IT_3B_G3.Cells(11, 5).Value = frmIT_3B_g3.cmb70.Text

            ' Friday

            Sheet_IT_3B_G3.Cells(2, 6).Value = frmIT_3B_g3.cmb81.Text
            Sheet_IT_3B_G3.Cells(3, 6).Value = frmIT_3B_g3.cmb82.Text
            Sheet_IT_3B_G3.Cells(4, 6).Value = frmIT_3B_g3.cmb83.Text
            Sheet_IT_3B_G3.Cells(5, 6).Value = frmIT_3B_g3.cmb84.Text
            Sheet_IT_3B_G3.Cells(6, 6).Value = frmIT_3B_g3.cmb85.Text
            Sheet_IT_3B_G3.Cells(7, 6).Value = frmIT_3B_g3.cmb86.Text
            Sheet_IT_3B_G3.Cells(8, 6).Value = frmIT_3B_g3.cmb87.Text
            Sheet_IT_3B_G3.Cells(9, 6).Value = frmIT_3B_g3.cmb88.Text
            Sheet_IT_3B_G3.Cells(10, 6).Value = frmIT_3B_g3.cmb89.Text
            Sheet_IT_3B_G3.Cells(11, 6).Value = frmIT_3B_g3.cmb90.Text
            '
            '
            '
            '
            '
            '
            '
            oWB = oXL.Workbooks.Add
            Sheet_IT_5A_G1 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_5A_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_5A_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_5A_G1.Cells(1, 1).Value = ""
            Sheet_IT_5A_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_5A_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_5A_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_5A_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_5A_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_5A_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_5A_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_5A_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_5A_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_5A_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_5A_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_5A_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_5A_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_5A_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_5A_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_5A_G1.Cells(2, 2).Value = frmIT_5A_G1.cmb1.Text
            Sheet_IT_5A_G1.Cells(3, 2).Value = frmIT_5A_G1.cmb2.Text
            Sheet_IT_5A_G1.Cells(4, 2).Value = frmIT_5A_G1.cmb3.Text
            Sheet_IT_5A_G1.Cells(5, 2).Value = frmIT_5A_G1.cmb4.Text
            Sheet_IT_5A_G1.Cells(6, 2).Value = frmIT_5A_G1.cmb5.Text
            Sheet_IT_5A_G1.Cells(7, 2).Value = frmIT_5A_G1.cmb6.Text
            Sheet_IT_5A_G1.Cells(8, 2).Value = frmIT_5A_G1.cmb7.Text
            Sheet_IT_5A_G1.Cells(9, 2).Value = frmIT_5A_G1.cmb8.Text
            Sheet_IT_5A_G1.Cells(10, 2).Value = frmIT_5A_G1.cmb9.Text
            Sheet_IT_5A_G1.Cells(11, 2).Value = frmIT_5A_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_5A_G1.Cells(2, 3).Value = frmIT_5A_G1.cmb21.Text
            Sheet_IT_5A_G1.Cells(3, 3).Value = frmIT_5A_G1.cmb22.Text
            Sheet_IT_5A_G1.Cells(4, 3).Value = frmIT_5A_G1.cmb23.Text
            Sheet_IT_5A_G1.Cells(5, 3).Value = frmIT_5A_G1.cmb24.Text
            Sheet_IT_5A_G1.Cells(6, 3).Value = frmIT_5A_G1.cmb25.Text
            Sheet_IT_5A_G1.Cells(7, 3).Value = frmIT_5A_G1.cmb26.Text
            Sheet_IT_5A_G1.Cells(8, 3).Value = frmIT_5A_G1.cmb27.Text
            Sheet_IT_5A_G1.Cells(9, 3).Value = frmIT_5A_G1.cmb28.Text
            Sheet_IT_5A_G1.Cells(10, 3).Value = frmIT_5A_G1.cmb29.Text
            Sheet_IT_5A_G1.Cells(11, 3).Value = frmIT_5A_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_5A_G1.Cells(2, 4).Value = frmIT_5A_G1.cmb41.Text
            Sheet_IT_5A_G1.Cells(3, 4).Value = frmIT_5A_G1.cmb42.Text
            Sheet_IT_5A_G1.Cells(4, 4).Value = frmIT_5A_G1.cmb43.Text
            Sheet_IT_5A_G1.Cells(5, 4).Value = frmIT_5A_G1.cmb44.Text
            Sheet_IT_5A_G1.Cells(6, 4).Value = frmIT_5A_G1.cmb45.Text
            Sheet_IT_5A_G1.Cells(7, 4).Value = frmIT_5A_G1.cmb46.Text
            Sheet_IT_5A_G1.Cells(8, 4).Value = frmIT_5A_G1.cmb47.Text
            Sheet_IT_5A_G1.Cells(9, 4).Value = frmIT_5A_G1.cmb48.Text
            Sheet_IT_5A_G1.Cells(10, 4).Value = frmIT_5A_G1.cmb49.Text
            Sheet_IT_5A_G1.Cells(11, 4).Value = frmIT_5A_G1.cmb50.Text

            ' Thursday

            Sheet_IT_5A_G1.Cells(2, 5).Value = frmIT_5A_G1.cmb61.Text
            Sheet_IT_5A_G1.Cells(3, 5).Value = frmIT_5A_G1.cmb62.Text
            Sheet_IT_5A_G1.Cells(4, 5).Value = frmIT_5A_G1.cmb63.Text
            Sheet_IT_5A_G1.Cells(5, 5).Value = frmIT_5A_G1.cmb64.Text
            Sheet_IT_5A_G1.Cells(6, 5).Value = frmIT_5A_G1.cmb65.Text
            Sheet_IT_5A_G1.Cells(7, 5).Value = frmIT_5A_G1.cmb66.Text
            Sheet_IT_5A_G1.Cells(8, 5).Value = frmIT_5A_G1.cmb67.Text
            Sheet_IT_5A_G1.Cells(9, 5).Value = frmIT_5A_G1.cmb68.Text
            Sheet_IT_5A_G1.Cells(10, 5).Value = frmIT_5A_G1.cmb69.Text
            Sheet_IT_5A_G1.Cells(11, 5).Value = frmIT_5A_G1.cmb70.Text

            ' Friday

            Sheet_IT_5A_G1.Cells(2, 6).Value = frmIT_5A_G1.cmb81.Text
            Sheet_IT_5A_G1.Cells(3, 6).Value = frmIT_5A_G1.cmb82.Text
            Sheet_IT_5A_G1.Cells(4, 6).Value = frmIT_5A_G1.cmb83.Text
            Sheet_IT_5A_G1.Cells(5, 6).Value = frmIT_5A_G1.cmb84.Text
            Sheet_IT_5A_G1.Cells(6, 6).Value = frmIT_5A_G1.cmb85.Text
            Sheet_IT_5A_G1.Cells(7, 6).Value = frmIT_5A_G1.cmb86.Text
            Sheet_IT_5A_G1.Cells(8, 6).Value = frmIT_5A_G1.cmb87.Text
            Sheet_IT_5A_G1.Cells(9, 6).Value = frmIT_5A_G1.cmb88.Text
            Sheet_IT_5A_G1.Cells(10, 6).Value = frmIT_5A_G1.cmb89.Text
            Sheet_IT_5A_G1.Cells(11, 6).Value = frmIT_5A_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''
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

            Sheet_IT_5A_G2.Cells(2, 2).Value = frmIT_5A_G2.cmb1.Text
            Sheet_IT_5A_G2.Cells(3, 2).Value = frmIT_5A_G2.cmb2.Text
            Sheet_IT_5A_G2.Cells(4, 2).Value = frmIT_5A_G2.cmb3.Text
            Sheet_IT_5A_G2.Cells(5, 2).Value = frmIT_5A_G2.cmb4.Text
            Sheet_IT_5A_G2.Cells(6, 2).Value = frmIT_5A_G2.cmb5.Text
            Sheet_IT_5A_G2.Cells(7, 2).Value = frmIT_5A_G2.cmb6.Text
            Sheet_IT_5A_G2.Cells(8, 2).Value = frmIT_5A_G2.cmb7.Text
            Sheet_IT_5A_G2.Cells(9, 2).Value = frmIT_5A_G2.cmb8.Text
            Sheet_IT_5A_G2.Cells(10, 2).Value = frmIT_5A_G2.cmb9.Text
            Sheet_IT_5A_G2.Cells(11, 2).Value = frmIT_5A_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_5A_G2.Cells(2, 3).Value = frmIT_5A_G2.cmb21.Text
            Sheet_IT_5A_G2.Cells(3, 3).Value = frmIT_5A_G2.cmb22.Text
            Sheet_IT_5A_G2.Cells(4, 3).Value = frmIT_5A_G2.cmb23.Text
            Sheet_IT_5A_G2.Cells(5, 3).Value = frmIT_5A_G2.cmb24.Text
            Sheet_IT_5A_G2.Cells(6, 3).Value = frmIT_5A_G2.cmb25.Text
            Sheet_IT_5A_G2.Cells(7, 3).Value = frmIT_5A_G2.cmb26.Text
            Sheet_IT_5A_G2.Cells(8, 3).Value = frmIT_5A_G2.cmb27.Text
            Sheet_IT_5A_G2.Cells(9, 3).Value = frmIT_5A_G2.cmb28.Text
            Sheet_IT_5A_G2.Cells(10, 3).Value = frmIT_5A_G2.cmb29.Text
            Sheet_IT_5A_G2.Cells(11, 3).Value = frmIT_5A_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_5A_G2.Cells(2, 4).Value = frmIT_5A_G2.cmb41.Text
            Sheet_IT_5A_G2.Cells(3, 4).Value = frmIT_5A_G2.cmb42.Text
            Sheet_IT_5A_G2.Cells(4, 4).Value = frmIT_5A_G2.cmb43.Text
            Sheet_IT_5A_G2.Cells(5, 4).Value = frmIT_5A_G2.cmb44.Text
            Sheet_IT_5A_G2.Cells(6, 4).Value = frmIT_5A_G2.cmb45.Text
            Sheet_IT_5A_G2.Cells(7, 4).Value = frmIT_5A_G2.cmb46.Text
            Sheet_IT_5A_G2.Cells(8, 4).Value = frmIT_5A_G2.cmb47.Text
            Sheet_IT_5A_G2.Cells(9, 4).Value = frmIT_5A_G2.cmb48.Text
            Sheet_IT_5A_G2.Cells(10, 4).Value = frmIT_5A_G2.cmb49.Text
            Sheet_IT_5A_G2.Cells(11, 4).Value = frmIT_5A_G2.cmb50.Text

            ' Thursday

            Sheet_IT_5A_G2.Cells(2, 5).Value = frmIT_5A_G2.cmb61.Text
            Sheet_IT_5A_G2.Cells(3, 5).Value = frmIT_5A_G2.cmb62.Text
            Sheet_IT_5A_G2.Cells(4, 5).Value = frmIT_5A_G2.cmb63.Text
            Sheet_IT_5A_G2.Cells(5, 5).Value = frmIT_5A_G2.cmb64.Text
            Sheet_IT_5A_G2.Cells(6, 5).Value = frmIT_5A_G2.cmb65.Text
            Sheet_IT_5A_G2.Cells(7, 5).Value = frmIT_5A_G2.cmb66.Text
            Sheet_IT_5A_G2.Cells(8, 5).Value = frmIT_5A_G2.cmb67.Text
            Sheet_IT_5A_G2.Cells(9, 5).Value = frmIT_5A_G2.cmb68.Text
            Sheet_IT_5A_G2.Cells(10, 5).Value = frmIT_5A_G2.cmb69.Text
            Sheet_IT_5A_G2.Cells(11, 5).Value = frmIT_5A_G2.cmb70.Text

            ' Friday

            Sheet_IT_5A_G2.Cells(2, 6).Value = frmIT_5A_G2.cmb81.Text
            Sheet_IT_5A_G2.Cells(3, 6).Value = frmIT_5A_G2.cmb82.Text
            Sheet_IT_5A_G2.Cells(4, 6).Value = frmIT_5A_G2.cmb83.Text
            Sheet_IT_5A_G2.Cells(5, 6).Value = frmIT_5A_G2.cmb84.Text
            Sheet_IT_5A_G2.Cells(6, 6).Value = frmIT_5A_G2.cmb85.Text
            Sheet_IT_5A_G2.Cells(7, 6).Value = frmIT_5A_G2.cmb86.Text
            Sheet_IT_5A_G2.Cells(8, 6).Value = frmIT_5A_G2.cmb87.Text
            Sheet_IT_5A_G2.Cells(9, 6).Value = frmIT_5A_G2.cmb88.Text
            Sheet_IT_5A_G2.Cells(10, 6).Value = frmIT_5A_G2.cmb89.Text
            Sheet_IT_5A_G2.Cells(11, 6).Value = frmIT_5A_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_5A_G3 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_5A_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_5A_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_5A_G3.Cells(1, 1).Value = ""
            Sheet_IT_5A_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_5A_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_5A_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_5A_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_5A_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_5A_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_5A_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_5A_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_5A_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_5A_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_5A_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_5A_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_5A_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_5A_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_5A_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_5A_G3.Cells(2, 2).Value = frmIT_5A_G3.cmb1.Text
            Sheet_IT_5A_G3.Cells(3, 2).Value = frmIT_5A_G3.cmb2.Text
            Sheet_IT_5A_G3.Cells(4, 2).Value = frmIT_5A_G3.cmb3.Text
            Sheet_IT_5A_G3.Cells(5, 2).Value = frmIT_5A_G3.cmb4.Text
            Sheet_IT_5A_G3.Cells(6, 2).Value = frmIT_5A_G3.cmb5.Text
            Sheet_IT_5A_G3.Cells(7, 2).Value = frmIT_5A_G3.cmb6.Text
            Sheet_IT_5A_G3.Cells(8, 2).Value = frmIT_5A_G3.cmb7.Text
            Sheet_IT_5A_G3.Cells(9, 2).Value = frmIT_5A_G3.cmb8.Text
            Sheet_IT_5A_G3.Cells(10, 2).Value = frmIT_5A_G3.cmb9.Text
            Sheet_IT_5A_G3.Cells(11, 2).Value = frmIT_5A_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_5A_G3.Cells(2, 3).Value = frmIT_5A_G3.cmb21.Text
            Sheet_IT_5A_G3.Cells(3, 3).Value = frmIT_5A_G3.cmb22.Text
            Sheet_IT_5A_G3.Cells(4, 3).Value = frmIT_5A_G3.cmb23.Text
            Sheet_IT_5A_G3.Cells(5, 3).Value = frmIT_5A_G3.cmb24.Text
            Sheet_IT_5A_G3.Cells(6, 3).Value = frmIT_5A_G3.cmb25.Text
            Sheet_IT_5A_G3.Cells(7, 3).Value = frmIT_5A_G3.cmb26.Text
            Sheet_IT_5A_G3.Cells(8, 3).Value = frmIT_5A_G3.cmb27.Text
            Sheet_IT_5A_G3.Cells(9, 3).Value = frmIT_5A_G3.cmb28.Text
            Sheet_IT_5A_G3.Cells(10, 3).Value = frmIT_5A_G3.cmb29.Text
            Sheet_IT_5A_G3.Cells(11, 3).Value = frmIT_5A_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_5A_G3.Cells(2, 4).Value = frmIT_5A_G3.cmb41.Text
            Sheet_IT_5A_G3.Cells(3, 4).Value = frmIT_5A_G3.cmb42.Text
            Sheet_IT_5A_G3.Cells(4, 4).Value = frmIT_5A_G3.cmb43.Text
            Sheet_IT_5A_G3.Cells(5, 4).Value = frmIT_5A_G3.cmb44.Text
            Sheet_IT_5A_G3.Cells(6, 4).Value = frmIT_5A_G3.cmb45.Text
            Sheet_IT_5A_G3.Cells(7, 4).Value = frmIT_5A_G3.cmb46.Text
            Sheet_IT_5A_G3.Cells(8, 4).Value = frmIT_5A_G3.cmb47.Text
            Sheet_IT_5A_G3.Cells(9, 4).Value = frmIT_5A_G3.cmb48.Text
            Sheet_IT_5A_G3.Cells(10, 4).Value = frmIT_5A_G3.cmb49.Text
            Sheet_IT_5A_G3.Cells(11, 4).Value = frmIT_5A_G3.cmb50.Text

            ' Thursday

            Sheet_IT_5A_G3.Cells(2, 5).Value = frmIT_5A_G3.cmb61.Text
            Sheet_IT_5A_G3.Cells(3, 5).Value = frmIT_5A_G3.cmb62.Text
            Sheet_IT_5A_G3.Cells(4, 5).Value = frmIT_5A_G3.cmb63.Text
            Sheet_IT_5A_G3.Cells(5, 5).Value = frmIT_5A_G3.cmb64.Text
            Sheet_IT_5A_G3.Cells(6, 5).Value = frmIT_5A_G3.cmb65.Text
            Sheet_IT_5A_G3.Cells(7, 5).Value = frmIT_5A_G3.cmb66.Text
            Sheet_IT_5A_G3.Cells(8, 5).Value = frmIT_5A_G3.cmb67.Text
            Sheet_IT_5A_G3.Cells(9, 5).Value = frmIT_5A_G3.cmb68.Text
            Sheet_IT_5A_G3.Cells(10, 5).Value = frmIT_5A_G3.cmb69.Text
            Sheet_IT_5A_G3.Cells(11, 5).Value = frmIT_5A_G3.cmb70.Text

            ' Friday

            Sheet_IT_5A_G3.Cells(2, 6).Value = frmIT_5A_G3.cmb81.Text
            Sheet_IT_5A_G3.Cells(3, 6).Value = frmIT_5A_G3.cmb82.Text
            Sheet_IT_5A_G3.Cells(4, 6).Value = frmIT_5A_G3.cmb83.Text
            Sheet_IT_5A_G3.Cells(5, 6).Value = frmIT_5A_G3.cmb84.Text
            Sheet_IT_5A_G3.Cells(6, 6).Value = frmIT_5A_G3.cmb85.Text
            Sheet_IT_5A_G3.Cells(7, 6).Value = frmIT_5A_G3.cmb86.Text
            Sheet_IT_5A_G3.Cells(8, 6).Value = frmIT_5A_G3.cmb87.Text
            Sheet_IT_5A_G3.Cells(9, 6).Value = frmIT_5A_G3.cmb88.Text
            Sheet_IT_5A_G3.Cells(10, 6).Value = frmIT_5A_G3.cmb89.Text
            Sheet_IT_5A_G3.Cells(11, 6).Value = frmIT_5A_G3.cmb90.Text
            ''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_5B_G1 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_5B_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_5B_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_5B_G1.Cells(1, 1).Value = ""
            Sheet_IT_5B_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_5B_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_5B_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_5B_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_5B_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_5B_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_5B_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_5B_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_5B_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_5B_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_5B_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_5B_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_5B_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_5B_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_5B_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_5B_G1.Cells(2, 2).Value = frmIT_5B_G1.cmb1.Text
            Sheet_IT_5B_G1.Cells(3, 2).Value = frmIT_5B_G1.cmb2.Text
            Sheet_IT_5B_G1.Cells(4, 2).Value = frmIT_5B_G1.cmb3.Text
            Sheet_IT_5B_G1.Cells(5, 2).Value = frmIT_5B_G1.cmb4.Text
            Sheet_IT_5B_G1.Cells(6, 2).Value = frmIT_5B_G1.cmb5.Text
            Sheet_IT_5B_G1.Cells(7, 2).Value = frmIT_5B_G1.cmb6.Text
            Sheet_IT_5B_G1.Cells(8, 2).Value = frmIT_5B_G1.cmb7.Text
            Sheet_IT_5B_G1.Cells(9, 2).Value = frmIT_5B_G1.cmb8.Text
            Sheet_IT_5B_G1.Cells(10, 2).Value = frmIT_5B_G1.cmb9.Text
            Sheet_IT_5B_G1.Cells(11, 2).Value = frmIT_5B_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_5B_G1.Cells(2, 3).Value = frmIT_5B_G1.cmb21.Text
            Sheet_IT_5B_G1.Cells(3, 3).Value = frmIT_5B_G1.cmb22.Text
            Sheet_IT_5B_G1.Cells(4, 3).Value = frmIT_5B_G1.cmb23.Text
            Sheet_IT_5B_G1.Cells(5, 3).Value = frmIT_5B_G1.cmb24.Text
            Sheet_IT_5B_G1.Cells(6, 3).Value = frmIT_5B_G1.cmb25.Text
            Sheet_IT_5B_G1.Cells(7, 3).Value = frmIT_5B_G1.cmb26.Text
            Sheet_IT_5B_G1.Cells(8, 3).Value = frmIT_5B_G1.cmb27.Text
            Sheet_IT_5B_G1.Cells(9, 3).Value = frmIT_5B_G1.cmb28.Text
            Sheet_IT_5B_G1.Cells(10, 3).Value = frmIT_5B_G1.cmb29.Text
            Sheet_IT_5B_G1.Cells(11, 3).Value = frmIT_5B_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_5B_G1.Cells(2, 4).Value = frmIT_5B_G1.cmb41.Text
            Sheet_IT_5B_G1.Cells(3, 4).Value = frmIT_5B_G1.cmb42.Text
            Sheet_IT_5B_G1.Cells(4, 4).Value = frmIT_5B_G1.cmb43.Text
            Sheet_IT_5B_G1.Cells(5, 4).Value = frmIT_5B_G1.cmb44.Text
            Sheet_IT_5B_G1.Cells(6, 4).Value = frmIT_5B_G1.cmb45.Text
            Sheet_IT_5B_G1.Cells(7, 4).Value = frmIT_5B_G1.cmb46.Text
            Sheet_IT_5B_G1.Cells(8, 4).Value = frmIT_5B_G1.cmb47.Text
            Sheet_IT_5B_G1.Cells(9, 4).Value = frmIT_5B_G1.cmb48.Text
            Sheet_IT_5B_G1.Cells(10, 4).Value = frmIT_5B_G1.cmb49.Text
            Sheet_IT_5B_G1.Cells(11, 4).Value = frmIT_5B_G1.cmb50.Text

            ' Thursday

            Sheet_IT_5B_G1.Cells(2, 5).Value = frmIT_5B_G1.cmb61.Text
            Sheet_IT_5B_G1.Cells(3, 5).Value = frmIT_5B_G1.cmb62.Text
            Sheet_IT_5B_G1.Cells(4, 5).Value = frmIT_5B_G1.cmb63.Text
            Sheet_IT_5B_G1.Cells(5, 5).Value = frmIT_5B_G1.cmb64.Text
            Sheet_IT_5B_G1.Cells(6, 5).Value = frmIT_5B_G1.cmb65.Text
            Sheet_IT_5B_G1.Cells(7, 5).Value = frmIT_5B_G1.cmb66.Text
            Sheet_IT_5B_G1.Cells(8, 5).Value = frmIT_5B_G1.cmb67.Text
            Sheet_IT_5B_G1.Cells(9, 5).Value = frmIT_5B_G1.cmb68.Text
            Sheet_IT_5B_G1.Cells(10, 5).Value = frmIT_5B_G1.cmb69.Text
            Sheet_IT_5B_G1.Cells(11, 5).Value = frmIT_5B_G1.cmb70.Text

            ' Friday

            Sheet_IT_5B_G1.Cells(2, 6).Value = frmIT_5B_G1.cmb81.Text
            Sheet_IT_5B_G1.Cells(3, 6).Value = frmIT_5B_G1.cmb82.Text
            Sheet_IT_5B_G1.Cells(4, 6).Value = frmIT_5B_G1.cmb83.Text
            Sheet_IT_5B_G1.Cells(5, 6).Value = frmIT_5B_G1.cmb84.Text
            Sheet_IT_5B_G1.Cells(6, 6).Value = frmIT_5B_G1.cmb85.Text
            Sheet_IT_5B_G1.Cells(7, 6).Value = frmIT_5B_G1.cmb86.Text
            Sheet_IT_5B_G1.Cells(8, 6).Value = frmIT_5B_G1.cmb87.Text
            Sheet_IT_5B_G1.Cells(9, 6).Value = frmIT_5B_G1.cmb88.Text
            Sheet_IT_5B_G1.Cells(10, 6).Value = frmIT_5B_G1.cmb89.Text
            Sheet_IT_5B_G1.Cells(11, 6).Value = frmIT_5B_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_5B_G2 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_5B_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_5B_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_5B_G2.Cells(1, 1).Value = ""
            Sheet_IT_5B_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_5B_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_5B_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_5B_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_5B_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_5B_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_5B_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_5B_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_5B_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_5B_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_5B_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_5B_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_5B_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_5B_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_5B_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_5B_G2.Cells(2, 2).Value = frmIT_5B_G2.cmb1.Text
            Sheet_IT_5B_G2.Cells(3, 2).Value = frmIT_5B_G2.cmb2.Text
            Sheet_IT_5B_G2.Cells(4, 2).Value = frmIT_5B_G2.cmb3.Text
            Sheet_IT_5B_G2.Cells(5, 2).Value = frmIT_5B_G2.cmb4.Text
            Sheet_IT_5B_G2.Cells(6, 2).Value = frmIT_5B_G2.cmb5.Text
            Sheet_IT_5B_G2.Cells(7, 2).Value = frmIT_5B_G2.cmb6.Text
            Sheet_IT_5B_G2.Cells(8, 2).Value = frmIT_5B_G2.cmb7.Text
            Sheet_IT_5B_G2.Cells(9, 2).Value = frmIT_5B_G2.cmb8.Text
            Sheet_IT_5B_G2.Cells(10, 2).Value = frmIT_5B_G2.cmb9.Text
            Sheet_IT_5B_G2.Cells(11, 2).Value = frmIT_5B_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_5B_G2.Cells(2, 3).Value = frmIT_5B_G2.cmb21.Text
            Sheet_IT_5B_G2.Cells(3, 3).Value = frmIT_5B_G2.cmb22.Text
            Sheet_IT_5B_G2.Cells(4, 3).Value = frmIT_5B_G2.cmb23.Text
            Sheet_IT_5B_G2.Cells(5, 3).Value = frmIT_5B_G2.cmb24.Text
            Sheet_IT_5B_G2.Cells(6, 3).Value = frmIT_5B_G2.cmb25.Text
            Sheet_IT_5B_G2.Cells(7, 3).Value = frmIT_5B_G2.cmb26.Text
            Sheet_IT_5B_G2.Cells(8, 3).Value = frmIT_5B_G2.cmb27.Text
            Sheet_IT_5B_G2.Cells(9, 3).Value = frmIT_5B_G2.cmb28.Text
            Sheet_IT_5B_G2.Cells(10, 3).Value = frmIT_5B_G2.cmb29.Text
            Sheet_IT_5B_G2.Cells(11, 3).Value = frmIT_5B_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_5B_G2.Cells(2, 4).Value = frmIT_5B_G2.cmb41.Text
            Sheet_IT_5B_G2.Cells(3, 4).Value = frmIT_5B_G2.cmb42.Text
            Sheet_IT_5B_G2.Cells(4, 4).Value = frmIT_5B_G2.cmb43.Text
            Sheet_IT_5B_G2.Cells(5, 4).Value = frmIT_5B_G2.cmb44.Text
            Sheet_IT_5B_G2.Cells(6, 4).Value = frmIT_5B_G2.cmb45.Text
            Sheet_IT_5B_G2.Cells(7, 4).Value = frmIT_5B_G2.cmb46.Text
            Sheet_IT_5B_G2.Cells(8, 4).Value = frmIT_5B_G2.cmb47.Text
            Sheet_IT_5B_G2.Cells(9, 4).Value = frmIT_5B_G2.cmb48.Text
            Sheet_IT_5B_G2.Cells(10, 4).Value = frmIT_5B_G2.cmb49.Text
            Sheet_IT_5B_G2.Cells(11, 4).Value = frmIT_5B_G2.cmb50.Text

            ' Thursday

            Sheet_IT_5B_G2.Cells(2, 5).Value = frmIT_5B_G2.cmb61.Text
            Sheet_IT_5B_G2.Cells(3, 5).Value = frmIT_5B_G2.cmb62.Text
            Sheet_IT_5B_G2.Cells(4, 5).Value = frmIT_5B_G2.cmb63.Text
            Sheet_IT_5B_G2.Cells(5, 5).Value = frmIT_5B_G2.cmb64.Text
            Sheet_IT_5B_G2.Cells(6, 5).Value = frmIT_5B_G2.cmb65.Text
            Sheet_IT_5B_G2.Cells(7, 5).Value = frmIT_5B_G2.cmb66.Text
            Sheet_IT_5B_G2.Cells(8, 5).Value = frmIT_5B_G2.cmb67.Text
            Sheet_IT_5B_G2.Cells(9, 5).Value = frmIT_5B_G2.cmb68.Text
            Sheet_IT_5B_G2.Cells(10, 5).Value = frmIT_5B_G2.cmb69.Text
            Sheet_IT_5B_G2.Cells(11, 5).Value = frmIT_5B_G2.cmb70.Text

            ' Friday

            Sheet_IT_5B_G2.Cells(2, 6).Value = frmIT_5B_G2.cmb81.Text
            Sheet_IT_5B_G2.Cells(3, 6).Value = frmIT_5B_G2.cmb82.Text
            Sheet_IT_5B_G2.Cells(4, 6).Value = frmIT_5B_G2.cmb83.Text
            Sheet_IT_5B_G2.Cells(5, 6).Value = frmIT_5B_G2.cmb84.Text
            Sheet_IT_5B_G2.Cells(6, 6).Value = frmIT_5B_G2.cmb85.Text
            Sheet_IT_5B_G2.Cells(7, 6).Value = frmIT_5B_G2.cmb86.Text
            Sheet_IT_5B_G2.Cells(8, 6).Value = frmIT_5B_G2.cmb87.Text
            Sheet_IT_5B_G2.Cells(9, 6).Value = frmIT_5B_G2.cmb88.Text
            Sheet_IT_5B_G2.Cells(10, 6).Value = frmIT_5B_G2.cmb89.Text
            Sheet_IT_5B_G2.Cells(11, 6).Value = frmIT_5B_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_5B_G3 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_5B_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_5B_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_5B_G3.Cells(1, 1).Value = ""
            Sheet_IT_5B_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_5B_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_5B_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_5B_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_5B_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_5B_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_5B_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_5B_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_5B_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_5B_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_5B_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_5B_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_5B_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_5B_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_5B_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_5B_G3.Cells(2, 2).Value = frmIT_5B_g3.cmb1.Text
            Sheet_IT_5B_G3.Cells(3, 2).Value = frmIT_5B_g3.cmb2.Text
            Sheet_IT_5B_G3.Cells(4, 2).Value = frmIT_5B_g3.cmb3.Text
            Sheet_IT_5B_G3.Cells(5, 2).Value = frmIT_5B_g3.cmb4.Text
            Sheet_IT_5B_G3.Cells(6, 2).Value = frmIT_5B_g3.cmb5.Text
            Sheet_IT_5B_G3.Cells(7, 2).Value = frmIT_5B_g3.cmb6.Text
            Sheet_IT_5B_G3.Cells(8, 2).Value = frmIT_5B_g3.cmb7.Text
            Sheet_IT_5B_G3.Cells(9, 2).Value = frmIT_5B_g3.cmb8.Text
            Sheet_IT_5B_G3.Cells(10, 2).Value = frmIT_5B_g3.cmb9.Text
            Sheet_IT_5B_G3.Cells(11, 2).Value = frmIT_5B_g3.cmb10.Text

            ' Tuesday

            Sheet_IT_5B_G3.Cells(2, 3).Value = frmIT_5B_g3.cmb21.Text
            Sheet_IT_5B_G3.Cells(3, 3).Value = frmIT_5B_g3.cmb22.Text
            Sheet_IT_5B_G3.Cells(4, 3).Value = frmIT_5B_g3.cmb23.Text
            Sheet_IT_5B_G3.Cells(5, 3).Value = frmIT_5B_g3.cmb24.Text
            Sheet_IT_5B_G3.Cells(6, 3).Value = frmIT_5B_g3.cmb25.Text
            Sheet_IT_5B_G3.Cells(7, 3).Value = frmIT_5B_g3.cmb26.Text
            Sheet_IT_5B_G3.Cells(8, 3).Value = frmIT_5B_g3.cmb27.Text
            Sheet_IT_5B_G3.Cells(9, 3).Value = frmIT_5B_g3.cmb28.Text
            Sheet_IT_5B_G3.Cells(10, 3).Value = frmIT_5B_g3.cmb29.Text
            Sheet_IT_5B_G3.Cells(11, 3).Value = frmIT_5B_g3.cmb30.Text

            ' Wednessday

            Sheet_IT_5B_G3.Cells(2, 4).Value = frmIT_5B_g3.cmb41.Text
            Sheet_IT_5B_G3.Cells(3, 4).Value = frmIT_5B_g3.cmb42.Text
            Sheet_IT_5B_G3.Cells(4, 4).Value = frmIT_5B_g3.cmb43.Text
            Sheet_IT_5B_G3.Cells(5, 4).Value = frmIT_5B_g3.cmb44.Text
            Sheet_IT_5B_G3.Cells(6, 4).Value = frmIT_5B_g3.cmb45.Text
            Sheet_IT_5B_G3.Cells(7, 4).Value = frmIT_5B_g3.cmb46.Text
            Sheet_IT_5B_G3.Cells(8, 4).Value = frmIT_5B_g3.cmb47.Text
            Sheet_IT_5B_G3.Cells(9, 4).Value = frmIT_5B_g3.cmb48.Text
            Sheet_IT_5B_G3.Cells(10, 4).Value = frmIT_5B_g3.cmb49.Text
            Sheet_IT_5B_G3.Cells(11, 4).Value = frmIT_5B_g3.cmb50.Text

            ' Thursday

            Sheet_IT_5B_G3.Cells(2, 5).Value = frmIT_5B_g3.cmb61.Text
            Sheet_IT_5B_G3.Cells(3, 5).Value = frmIT_5B_g3.cmb62.Text
            Sheet_IT_5B_G3.Cells(4, 5).Value = frmIT_5B_g3.cmb63.Text
            Sheet_IT_5B_G3.Cells(5, 5).Value = frmIT_5B_g3.cmb64.Text
            Sheet_IT_5B_G3.Cells(6, 5).Value = frmIT_5B_g3.cmb65.Text
            Sheet_IT_5B_G3.Cells(7, 5).Value = frmIT_5B_g3.cmb66.Text
            Sheet_IT_5B_G3.Cells(8, 5).Value = frmIT_5B_g3.cmb67.Text
            Sheet_IT_5B_G3.Cells(9, 5).Value = frmIT_5B_g3.cmb68.Text
            Sheet_IT_5B_G3.Cells(10, 5).Value = frmIT_5B_g3.cmb69.Text
            Sheet_IT_5B_G3.Cells(11, 5).Value = frmIT_5B_g3.cmb70.Text

            ' Friday

            Sheet_IT_5B_G3.Cells(2, 6).Value = frmIT_5B_g3.cmb81.Text
            Sheet_IT_5B_G3.Cells(3, 6).Value = frmIT_5B_g3.cmb82.Text
            Sheet_IT_5B_G3.Cells(4, 6).Value = frmIT_5B_g3.cmb83.Text
            Sheet_IT_5B_G3.Cells(5, 6).Value = frmIT_5B_g3.cmb84.Text
            Sheet_IT_5B_G3.Cells(6, 6).Value = frmIT_5B_g3.cmb85.Text
            Sheet_IT_5B_G3.Cells(7, 6).Value = frmIT_5B_g3.cmb86.Text
            Sheet_IT_5B_G3.Cells(8, 6).Value = frmIT_5B_g3.cmb87.Text
            Sheet_IT_5B_G3.Cells(9, 6).Value = frmIT_5B_g3.cmb88.Text
            Sheet_IT_5B_G3.Cells(10, 6).Value = frmIT_5B_g3.cmb89.Text
            Sheet_IT_5B_G3.Cells(11, 6).Value = frmIT_5B_g3.cmb90.Text

            '
            '
            '
            '
            '
            '
            '
            oWB = oXL.Workbooks.Add
            Sheet_IT_7A_G1 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_7A_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_7A_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_7A_G1.Cells(1, 1).Value = ""
            Sheet_IT_7A_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_7A_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_7A_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_7A_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_7A_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_7A_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_7A_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_7A_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_7A_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_7A_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_7A_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_7A_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_7A_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_7A_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_7A_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_7A_G1.Cells(2, 2).Value = frmIT_7A_G1.cmb1.Text
            Sheet_IT_7A_G1.Cells(3, 2).Value = frmIT_7A_G1.cmb2.Text
            Sheet_IT_7A_G1.Cells(4, 2).Value = frmIT_7A_G1.cmb3.Text
            Sheet_IT_7A_G1.Cells(5, 2).Value = frmIT_7A_G1.cmb4.Text
            Sheet_IT_7A_G1.Cells(6, 2).Value = frmIT_7A_G1.cmb5.Text
            Sheet_IT_7A_G1.Cells(7, 2).Value = frmIT_7A_G1.cmb6.Text
            Sheet_IT_7A_G1.Cells(8, 2).Value = frmIT_7A_G1.cmb7.Text
            Sheet_IT_7A_G1.Cells(9, 2).Value = frmIT_7A_G1.cmb8.Text
            Sheet_IT_7A_G1.Cells(10, 2).Value = frmIT_7A_G1.cmb9.Text
            Sheet_IT_7A_G1.Cells(11, 2).Value = frmIT_7A_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_7A_G1.Cells(2, 3).Value = frmIT_7A_G1.cmb21.Text
            Sheet_IT_7A_G1.Cells(3, 3).Value = frmIT_7A_G1.cmb22.Text
            Sheet_IT_7A_G1.Cells(4, 3).Value = frmIT_7A_G1.cmb23.Text
            Sheet_IT_7A_G1.Cells(5, 3).Value = frmIT_7A_G1.cmb24.Text
            Sheet_IT_7A_G1.Cells(6, 3).Value = frmIT_7A_G1.cmb25.Text
            Sheet_IT_7A_G1.Cells(7, 3).Value = frmIT_7A_G1.cmb26.Text
            Sheet_IT_7A_G1.Cells(8, 3).Value = frmIT_7A_G1.cmb27.Text
            Sheet_IT_7A_G1.Cells(9, 3).Value = frmIT_7A_G1.cmb28.Text
            Sheet_IT_7A_G1.Cells(10, 3).Value = frmIT_7A_G1.cmb29.Text
            Sheet_IT_7A_G1.Cells(11, 3).Value = frmIT_7A_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_7A_G1.Cells(2, 4).Value = frmIT_7A_G1.cmb41.Text
            Sheet_IT_7A_G1.Cells(3, 4).Value = frmIT_7A_G1.cmb42.Text
            Sheet_IT_7A_G1.Cells(4, 4).Value = frmIT_7A_G1.cmb43.Text
            Sheet_IT_7A_G1.Cells(5, 4).Value = frmIT_7A_G1.cmb44.Text
            Sheet_IT_7A_G1.Cells(6, 4).Value = frmIT_7A_G1.cmb45.Text
            Sheet_IT_7A_G1.Cells(7, 4).Value = frmIT_7A_G1.cmb46.Text
            Sheet_IT_7A_G1.Cells(8, 4).Value = frmIT_7A_G1.cmb47.Text
            Sheet_IT_7A_G1.Cells(9, 4).Value = frmIT_7A_G1.cmb48.Text
            Sheet_IT_7A_G1.Cells(10, 4).Value = frmIT_7A_G1.cmb49.Text
            Sheet_IT_7A_G1.Cells(11, 4).Value = frmIT_7A_G1.cmb50.Text

            ' Thursday

            Sheet_IT_7A_G1.Cells(2, 5).Value = frmIT_7A_G1.cmb61.Text
            Sheet_IT_7A_G1.Cells(3, 5).Value = frmIT_7A_G1.cmb62.Text
            Sheet_IT_7A_G1.Cells(4, 5).Value = frmIT_7A_G1.cmb63.Text
            Sheet_IT_7A_G1.Cells(5, 5).Value = frmIT_7A_G1.cmb64.Text
            Sheet_IT_7A_G1.Cells(6, 5).Value = frmIT_7A_G1.cmb65.Text
            Sheet_IT_7A_G1.Cells(7, 5).Value = frmIT_7A_G1.cmb66.Text
            Sheet_IT_7A_G1.Cells(8, 5).Value = frmIT_7A_G1.cmb67.Text
            Sheet_IT_7A_G1.Cells(9, 5).Value = frmIT_7A_G1.cmb68.Text
            Sheet_IT_7A_G1.Cells(10, 5).Value = frmIT_7A_G1.cmb69.Text
            Sheet_IT_7A_G1.Cells(11, 5).Value = frmIT_7A_G1.cmb70.Text

            ' Friday

            Sheet_IT_7A_G1.Cells(2, 6).Value = frmIT_7A_G1.cmb81.Text
            Sheet_IT_7A_G1.Cells(3, 6).Value = frmIT_7A_G1.cmb82.Text
            Sheet_IT_7A_G1.Cells(4, 6).Value = frmIT_7A_G1.cmb83.Text
            Sheet_IT_7A_G1.Cells(5, 6).Value = frmIT_7A_G1.cmb84.Text
            Sheet_IT_7A_G1.Cells(6, 6).Value = frmIT_7A_G1.cmb85.Text
            Sheet_IT_7A_G1.Cells(7, 6).Value = frmIT_7A_G1.cmb86.Text
            Sheet_IT_7A_G1.Cells(8, 6).Value = frmIT_7A_G1.cmb87.Text
            Sheet_IT_7A_G1.Cells(9, 6).Value = frmIT_7A_G1.cmb88.Text
            Sheet_IT_7A_G1.Cells(10, 6).Value = frmIT_7A_G1.cmb89.Text
            Sheet_IT_7A_G1.Cells(11, 6).Value = frmIT_7A_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_7A_G2 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_7A_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_7A_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_7A_G2.Cells(1, 1).Value = ""
            Sheet_IT_7A_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_7A_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_7A_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_7A_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_7A_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_7A_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_7A_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_7A_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_7A_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_7A_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_7A_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_7A_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_7A_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_7A_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_7A_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_7A_G2.Cells(2, 2).Value = frmIT_7A_G2.cmb1.Text
            Sheet_IT_7A_G2.Cells(3, 2).Value = frmIT_7A_G2.cmb2.Text
            Sheet_IT_7A_G2.Cells(4, 2).Value = frmIT_7A_G2.cmb3.Text
            Sheet_IT_7A_G2.Cells(5, 2).Value = frmIT_7A_G2.cmb4.Text
            Sheet_IT_7A_G2.Cells(6, 2).Value = frmIT_7A_G2.cmb5.Text
            Sheet_IT_7A_G2.Cells(7, 2).Value = frmIT_7A_G2.cmb6.Text
            Sheet_IT_7A_G2.Cells(8, 2).Value = frmIT_7A_G2.cmb7.Text
            Sheet_IT_7A_G2.Cells(9, 2).Value = frmIT_7A_G2.cmb8.Text
            Sheet_IT_7A_G2.Cells(10, 2).Value = frmIT_7A_G2.cmb9.Text
            Sheet_IT_7A_G2.Cells(11, 2).Value = frmIT_7A_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_7A_G2.Cells(2, 3).Value = frmIT_7A_G2.cmb21.Text
            Sheet_IT_7A_G2.Cells(3, 3).Value = frmIT_7A_G2.cmb22.Text
            Sheet_IT_7A_G2.Cells(4, 3).Value = frmIT_7A_G2.cmb23.Text
            Sheet_IT_7A_G2.Cells(5, 3).Value = frmIT_7A_G2.cmb24.Text
            Sheet_IT_7A_G2.Cells(6, 3).Value = frmIT_7A_G2.cmb25.Text
            Sheet_IT_7A_G2.Cells(7, 3).Value = frmIT_7A_G2.cmb26.Text
            Sheet_IT_7A_G2.Cells(8, 3).Value = frmIT_7A_G2.cmb27.Text
            Sheet_IT_7A_G2.Cells(9, 3).Value = frmIT_7A_G2.cmb28.Text
            Sheet_IT_7A_G2.Cells(10, 3).Value = frmIT_7A_G2.cmb29.Text
            Sheet_IT_7A_G2.Cells(11, 3).Value = frmIT_7A_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_7A_G2.Cells(2, 4).Value = frmIT_7A_G2.cmb41.Text
            Sheet_IT_7A_G2.Cells(3, 4).Value = frmIT_7A_G2.cmb42.Text
            Sheet_IT_7A_G2.Cells(4, 4).Value = frmIT_7A_G2.cmb43.Text
            Sheet_IT_7A_G2.Cells(5, 4).Value = frmIT_7A_G2.cmb44.Text
            Sheet_IT_7A_G2.Cells(6, 4).Value = frmIT_7A_G2.cmb45.Text
            Sheet_IT_7A_G2.Cells(7, 4).Value = frmIT_7A_G2.cmb46.Text
            Sheet_IT_7A_G2.Cells(8, 4).Value = frmIT_7A_G2.cmb47.Text
            Sheet_IT_7A_G2.Cells(9, 4).Value = frmIT_7A_G2.cmb48.Text
            Sheet_IT_7A_G2.Cells(10, 4).Value = frmIT_7A_G2.cmb49.Text
            Sheet_IT_7A_G2.Cells(11, 4).Value = frmIT_7A_G2.cmb50.Text

            ' Thursday

            Sheet_IT_7A_G2.Cells(2, 5).Value = frmIT_7A_G2.cmb61.Text
            Sheet_IT_7A_G2.Cells(3, 5).Value = frmIT_7A_G2.cmb62.Text
            Sheet_IT_7A_G2.Cells(4, 5).Value = frmIT_7A_G2.cmb63.Text
            Sheet_IT_7A_G2.Cells(5, 5).Value = frmIT_7A_G2.cmb64.Text
            Sheet_IT_7A_G2.Cells(6, 5).Value = frmIT_7A_G2.cmb65.Text
            Sheet_IT_7A_G2.Cells(7, 5).Value = frmIT_7A_G2.cmb66.Text
            Sheet_IT_7A_G2.Cells(8, 5).Value = frmIT_7A_G2.cmb67.Text
            Sheet_IT_7A_G2.Cells(9, 5).Value = frmIT_7A_G2.cmb68.Text
            Sheet_IT_7A_G2.Cells(10, 5).Value = frmIT_7A_G2.cmb69.Text
            Sheet_IT_7A_G2.Cells(11, 5).Value = frmIT_7A_G2.cmb70.Text

            ' Friday

            Sheet_IT_7A_G2.Cells(2, 6).Value = frmIT_7A_G2.cmb81.Text
            Sheet_IT_7A_G2.Cells(3, 6).Value = frmIT_7A_G2.cmb82.Text
            Sheet_IT_7A_G2.Cells(4, 6).Value = frmIT_7A_G2.cmb83.Text
            Sheet_IT_7A_G2.Cells(5, 6).Value = frmIT_7A_G2.cmb84.Text
            Sheet_IT_7A_G2.Cells(6, 6).Value = frmIT_7A_G2.cmb85.Text
            Sheet_IT_7A_G2.Cells(7, 6).Value = frmIT_7A_G2.cmb86.Text
            Sheet_IT_7A_G2.Cells(8, 6).Value = frmIT_7A_G2.cmb87.Text
            Sheet_IT_7A_G2.Cells(9, 6).Value = frmIT_7A_G2.cmb88.Text
            Sheet_IT_7A_G2.Cells(10, 6).Value = frmIT_7A_G2.cmb89.Text
            Sheet_IT_7A_G2.Cells(11, 6).Value = frmIT_7A_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_7A_G3 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_7A_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_7A_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_7A_G3.Cells(1, 1).Value = ""
            Sheet_IT_7A_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_7A_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_7A_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_7A_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_7A_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_7A_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_7A_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_7A_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_7A_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_7A_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_7A_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_7A_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_7A_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_7A_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_7A_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_7A_G3.Cells(2, 2).Value = frmIT_7A_G3.cmb1.Text
            Sheet_IT_7A_G3.Cells(3, 2).Value = frmIT_7A_G3.cmb2.Text
            Sheet_IT_7A_G3.Cells(4, 2).Value = frmIT_7A_G3.cmb3.Text
            Sheet_IT_7A_G3.Cells(5, 2).Value = frmIT_7A_G3.cmb4.Text
            Sheet_IT_7A_G3.Cells(6, 2).Value = frmIT_7A_G3.cmb5.Text
            Sheet_IT_7A_G3.Cells(7, 2).Value = frmIT_7A_G3.cmb6.Text
            Sheet_IT_7A_G3.Cells(8, 2).Value = frmIT_7A_G3.cmb7.Text
            Sheet_IT_7A_G3.Cells(9, 2).Value = frmIT_7A_G3.cmb8.Text
            Sheet_IT_7A_G3.Cells(10, 2).Value = frmIT_7A_G3.cmb9.Text
            Sheet_IT_7A_G3.Cells(11, 2).Value = frmIT_7A_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_7A_G3.Cells(2, 3).Value = frmIT_7A_G3.cmb21.Text
            Sheet_IT_7A_G3.Cells(3, 3).Value = frmIT_7A_G3.cmb22.Text
            Sheet_IT_7A_G3.Cells(4, 3).Value = frmIT_7A_G3.cmb23.Text
            Sheet_IT_7A_G3.Cells(5, 3).Value = frmIT_7A_G3.cmb24.Text
            Sheet_IT_7A_G3.Cells(6, 3).Value = frmIT_7A_G3.cmb25.Text
            Sheet_IT_7A_G3.Cells(7, 3).Value = frmIT_7A_G3.cmb26.Text
            Sheet_IT_7A_G3.Cells(8, 3).Value = frmIT_7A_G3.cmb27.Text
            Sheet_IT_7A_G3.Cells(9, 3).Value = frmIT_7A_G3.cmb28.Text
            Sheet_IT_7A_G3.Cells(10, 3).Value = frmIT_7A_G3.cmb29.Text
            Sheet_IT_7A_G3.Cells(11, 3).Value = frmIT_7A_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_7A_G3.Cells(2, 4).Value = frmIT_7A_G3.cmb41.Text
            Sheet_IT_7A_G3.Cells(3, 4).Value = frmIT_7A_G3.cmb42.Text
            Sheet_IT_7A_G3.Cells(4, 4).Value = frmIT_7A_G3.cmb43.Text
            Sheet_IT_7A_G3.Cells(5, 4).Value = frmIT_7A_G3.cmb44.Text
            Sheet_IT_7A_G3.Cells(6, 4).Value = frmIT_7A_G3.cmb45.Text
            Sheet_IT_7A_G3.Cells(7, 4).Value = frmIT_7A_G3.cmb46.Text
            Sheet_IT_7A_G3.Cells(8, 4).Value = frmIT_7A_G3.cmb47.Text
            Sheet_IT_7A_G3.Cells(9, 4).Value = frmIT_7A_G3.cmb48.Text
            Sheet_IT_7A_G3.Cells(10, 4).Value = frmIT_7A_G3.cmb49.Text
            Sheet_IT_7A_G3.Cells(11, 4).Value = frmIT_7A_G3.cmb50.Text

            ' Thursday

            Sheet_IT_7A_G3.Cells(2, 5).Value = frmIT_7A_G3.cmb61.Text
            Sheet_IT_7A_G3.Cells(3, 5).Value = frmIT_7A_G3.cmb62.Text
            Sheet_IT_7A_G3.Cells(4, 5).Value = frmIT_7A_G3.cmb63.Text
            Sheet_IT_7A_G3.Cells(5, 5).Value = frmIT_7A_G3.cmb64.Text
            Sheet_IT_7A_G3.Cells(6, 5).Value = frmIT_7A_G3.cmb65.Text
            Sheet_IT_7A_G3.Cells(7, 5).Value = frmIT_7A_G3.cmb66.Text
            Sheet_IT_7A_G3.Cells(8, 5).Value = frmIT_7A_G3.cmb67.Text
            Sheet_IT_7A_G3.Cells(9, 5).Value = frmIT_7A_G3.cmb68.Text
            Sheet_IT_7A_G3.Cells(10, 5).Value = frmIT_7A_G3.cmb69.Text
            Sheet_IT_7A_G3.Cells(11, 5).Value = frmIT_7A_G3.cmb70.Text

            ' Friday

            Sheet_IT_7A_G3.Cells(2, 6).Value = frmIT_7A_G3.cmb81.Text
            Sheet_IT_7A_G3.Cells(3, 6).Value = frmIT_7A_G3.cmb82.Text
            Sheet_IT_7A_G3.Cells(4, 6).Value = frmIT_7A_G3.cmb83.Text
            Sheet_IT_7A_G3.Cells(5, 6).Value = frmIT_7A_G3.cmb84.Text
            Sheet_IT_7A_G3.Cells(6, 6).Value = frmIT_7A_G3.cmb85.Text
            Sheet_IT_7A_G3.Cells(7, 6).Value = frmIT_7A_G3.cmb86.Text
            Sheet_IT_7A_G3.Cells(8, 6).Value = frmIT_7A_G3.cmb87.Text
            Sheet_IT_7A_G3.Cells(9, 6).Value = frmIT_7A_G3.cmb88.Text
            Sheet_IT_7A_G3.Cells(10, 6).Value = frmIT_7A_G3.cmb89.Text
            Sheet_IT_7A_G3.Cells(11, 6).Value = frmIT_7A_G3.cmb90.Text
            ''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_7B_G1 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_7B_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_7B_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_7B_G1.Cells(1, 1).Value = ""
            Sheet_IT_7B_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_7B_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_7B_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_7B_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_7B_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_7B_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_7B_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_7B_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_7B_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_7B_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_7B_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_7B_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_7B_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_7B_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_7B_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_7B_G1.Cells(2, 2).Value = frmIT_7B_G1.cmb1.Text
            Sheet_IT_7B_G1.Cells(3, 2).Value = frmIT_7B_G1.cmb2.Text
            Sheet_IT_7B_G1.Cells(4, 2).Value = frmIT_7B_G1.cmb3.Text
            Sheet_IT_7B_G1.Cells(5, 2).Value = frmIT_7B_G1.cmb4.Text
            Sheet_IT_7B_G1.Cells(6, 2).Value = frmIT_7B_G1.cmb5.Text
            Sheet_IT_7B_G1.Cells(7, 2).Value = frmIT_7B_G1.cmb6.Text
            Sheet_IT_7B_G1.Cells(8, 2).Value = frmIT_7B_G1.cmb7.Text
            Sheet_IT_7B_G1.Cells(9, 2).Value = frmIT_7B_G1.cmb8.Text
            Sheet_IT_7B_G1.Cells(10, 2).Value = frmIT_7B_G1.cmb9.Text
            Sheet_IT_7B_G1.Cells(11, 2).Value = frmIT_7B_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_7B_G1.Cells(2, 3).Value = frmIT_7B_G1.cmb21.Text
            Sheet_IT_7B_G1.Cells(3, 3).Value = frmIT_7B_G1.cmb22.Text
            Sheet_IT_7B_G1.Cells(4, 3).Value = frmIT_7B_G1.cmb23.Text
            Sheet_IT_7B_G1.Cells(5, 3).Value = frmIT_7B_G1.cmb24.Text
            Sheet_IT_7B_G1.Cells(6, 3).Value = frmIT_7B_G1.cmb25.Text
            Sheet_IT_7B_G1.Cells(7, 3).Value = frmIT_7B_G1.cmb26.Text
            Sheet_IT_7B_G1.Cells(8, 3).Value = frmIT_7B_G1.cmb27.Text
            Sheet_IT_7B_G1.Cells(9, 3).Value = frmIT_7B_G1.cmb28.Text
            Sheet_IT_7B_G1.Cells(10, 3).Value = frmIT_7B_G1.cmb29.Text
            Sheet_IT_7B_G1.Cells(11, 3).Value = frmIT_7B_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_7B_G1.Cells(2, 4).Value = frmIT_7B_G1.cmb41.Text
            Sheet_IT_7B_G1.Cells(3, 4).Value = frmIT_7B_G1.cmb42.Text
            Sheet_IT_7B_G1.Cells(4, 4).Value = frmIT_7B_G1.cmb43.Text
            Sheet_IT_7B_G1.Cells(5, 4).Value = frmIT_7B_G1.cmb44.Text
            Sheet_IT_7B_G1.Cells(6, 4).Value = frmIT_7B_G1.cmb45.Text
            Sheet_IT_7B_G1.Cells(7, 4).Value = frmIT_7B_G1.cmb46.Text
            Sheet_IT_7B_G1.Cells(8, 4).Value = frmIT_7B_G1.cmb47.Text
            Sheet_IT_7B_G1.Cells(9, 4).Value = frmIT_7B_G1.cmb48.Text
            Sheet_IT_7B_G1.Cells(10, 4).Value = frmIT_7B_G1.cmb49.Text
            Sheet_IT_7B_G1.Cells(11, 4).Value = frmIT_7B_G1.cmb50.Text

            ' Thursday

            Sheet_IT_7B_G1.Cells(2, 5).Value = frmIT_7B_G1.cmb61.Text
            Sheet_IT_7B_G1.Cells(3, 5).Value = frmIT_7B_G1.cmb62.Text
            Sheet_IT_7B_G1.Cells(4, 5).Value = frmIT_7B_G1.cmb63.Text
            Sheet_IT_7B_G1.Cells(5, 5).Value = frmIT_7B_G1.cmb64.Text
            Sheet_IT_7B_G1.Cells(6, 5).Value = frmIT_7B_G1.cmb65.Text
            Sheet_IT_7B_G1.Cells(7, 5).Value = frmIT_7B_G1.cmb66.Text
            Sheet_IT_7B_G1.Cells(8, 5).Value = frmIT_7B_G1.cmb67.Text
            Sheet_IT_7B_G1.Cells(9, 5).Value = frmIT_7B_G1.cmb68.Text
            Sheet_IT_7B_G1.Cells(10, 5).Value = frmIT_7B_G1.cmb69.Text
            Sheet_IT_7B_G1.Cells(11, 5).Value = frmIT_7B_G1.cmb70.Text

            ' Friday

            Sheet_IT_7B_G1.Cells(2, 6).Value = frmIT_7B_G1.cmb81.Text
            Sheet_IT_7B_G1.Cells(3, 6).Value = frmIT_7B_G1.cmb82.Text
            Sheet_IT_7B_G1.Cells(4, 6).Value = frmIT_7B_G1.cmb83.Text
            Sheet_IT_7B_G1.Cells(5, 6).Value = frmIT_7B_G1.cmb84.Text
            Sheet_IT_7B_G1.Cells(6, 6).Value = frmIT_7B_G1.cmb85.Text
            Sheet_IT_7B_G1.Cells(7, 6).Value = frmIT_7B_G1.cmb86.Text
            Sheet_IT_7B_G1.Cells(8, 6).Value = frmIT_7B_G1.cmb87.Text
            Sheet_IT_7B_G1.Cells(9, 6).Value = frmIT_7B_G1.cmb88.Text
            Sheet_IT_7B_G1.Cells(10, 6).Value = frmIT_7B_G1.cmb89.Text
            Sheet_IT_7B_G1.Cells(11, 6).Value = frmIT_7B_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_7B_G2 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_7B_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_7B_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_7B_G2.Cells(1, 1).Value = ""
            Sheet_IT_7B_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_7B_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_7B_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_7B_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_7B_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_7B_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_7B_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_7B_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_7B_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_7B_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_7B_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_7B_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_7B_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_7B_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_7B_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_7B_G2.Cells(2, 2).Value = frmIT_7B_G2.cmb1.Text
            Sheet_IT_7B_G2.Cells(3, 2).Value = frmIT_7B_G2.cmb2.Text
            Sheet_IT_7B_G2.Cells(4, 2).Value = frmIT_7B_G2.cmb3.Text
            Sheet_IT_7B_G2.Cells(5, 2).Value = frmIT_7B_G2.cmb4.Text
            Sheet_IT_7B_G2.Cells(6, 2).Value = frmIT_7B_G2.cmb5.Text
            Sheet_IT_7B_G2.Cells(7, 2).Value = frmIT_7B_G2.cmb6.Text
            Sheet_IT_7B_G2.Cells(8, 2).Value = frmIT_7B_G2.cmb7.Text
            Sheet_IT_7B_G2.Cells(9, 2).Value = frmIT_7B_G2.cmb8.Text
            Sheet_IT_7B_G2.Cells(10, 2).Value = frmIT_7B_G2.cmb9.Text
            Sheet_IT_7B_G2.Cells(11, 2).Value = frmIT_7B_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_7B_G2.Cells(2, 3).Value = frmIT_7B_G2.cmb21.Text
            Sheet_IT_7B_G2.Cells(3, 3).Value = frmIT_7B_G2.cmb22.Text
            Sheet_IT_7B_G2.Cells(4, 3).Value = frmIT_7B_G2.cmb23.Text
            Sheet_IT_7B_G2.Cells(5, 3).Value = frmIT_7B_G2.cmb24.Text
            Sheet_IT_7B_G2.Cells(6, 3).Value = frmIT_7B_G2.cmb25.Text
            Sheet_IT_7B_G2.Cells(7, 3).Value = frmIT_7B_G2.cmb26.Text
            Sheet_IT_7B_G2.Cells(8, 3).Value = frmIT_7B_G2.cmb27.Text
            Sheet_IT_7B_G2.Cells(9, 3).Value = frmIT_7B_G2.cmb28.Text
            Sheet_IT_7B_G2.Cells(10, 3).Value = frmIT_7B_G2.cmb29.Text
            Sheet_IT_7B_G2.Cells(11, 3).Value = frmIT_7B_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_7B_G2.Cells(2, 4).Value = frmIT_7B_G2.cmb41.Text
            Sheet_IT_7B_G2.Cells(3, 4).Value = frmIT_7B_G2.cmb42.Text
            Sheet_IT_7B_G2.Cells(4, 4).Value = frmIT_7B_G2.cmb43.Text
            Sheet_IT_7B_G2.Cells(5, 4).Value = frmIT_7B_G2.cmb44.Text
            Sheet_IT_7B_G2.Cells(6, 4).Value = frmIT_7B_G2.cmb45.Text
            Sheet_IT_7B_G2.Cells(7, 4).Value = frmIT_7B_G2.cmb46.Text
            Sheet_IT_7B_G2.Cells(8, 4).Value = frmIT_7B_G2.cmb47.Text
            Sheet_IT_7B_G2.Cells(9, 4).Value = frmIT_7B_G2.cmb48.Text
            Sheet_IT_7B_G2.Cells(10, 4).Value = frmIT_7B_G2.cmb49.Text
            Sheet_IT_7B_G2.Cells(11, 4).Value = frmIT_7B_G2.cmb50.Text

            ' Thursday

            Sheet_IT_7B_G2.Cells(2, 5).Value = frmIT_7B_G2.cmb61.Text
            Sheet_IT_7B_G2.Cells(3, 5).Value = frmIT_7B_G2.cmb62.Text
            Sheet_IT_7B_G2.Cells(4, 5).Value = frmIT_7B_G2.cmb63.Text
            Sheet_IT_7B_G2.Cells(5, 5).Value = frmIT_7B_G2.cmb64.Text
            Sheet_IT_7B_G2.Cells(6, 5).Value = frmIT_7B_G2.cmb65.Text
            Sheet_IT_7B_G2.Cells(7, 5).Value = frmIT_7B_G2.cmb66.Text
            Sheet_IT_7B_G2.Cells(8, 5).Value = frmIT_7B_G2.cmb67.Text
            Sheet_IT_7B_G2.Cells(9, 5).Value = frmIT_7B_G2.cmb68.Text
            Sheet_IT_7B_G2.Cells(10, 5).Value = frmIT_7B_G2.cmb69.Text
            Sheet_IT_7B_G2.Cells(11, 5).Value = frmIT_7B_G2.cmb70.Text

            ' Friday

            Sheet_IT_7B_G2.Cells(2, 6).Value = frmIT_7B_G2.cmb81.Text
            Sheet_IT_7B_G2.Cells(3, 6).Value = frmIT_7B_G2.cmb82.Text
            Sheet_IT_7B_G2.Cells(4, 6).Value = frmIT_7B_G2.cmb83.Text
            Sheet_IT_7B_G2.Cells(5, 6).Value = frmIT_7B_G2.cmb84.Text
            Sheet_IT_7B_G2.Cells(6, 6).Value = frmIT_7B_G2.cmb85.Text
            Sheet_IT_7B_G2.Cells(7, 6).Value = frmIT_7B_G2.cmb86.Text
            Sheet_IT_7B_G2.Cells(8, 6).Value = frmIT_7B_G2.cmb87.Text
            Sheet_IT_7B_G2.Cells(9, 6).Value = frmIT_7B_G2.cmb88.Text
            Sheet_IT_7B_G2.Cells(10, 6).Value = frmIT_7B_G2.cmb89.Text
            Sheet_IT_7B_G2.Cells(11, 6).Value = frmIT_7B_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''
            oWB = oXL.Workbooks.Add
            Sheet_IT_7B_G3 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_7B_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_7B_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_7B_G3.Cells(1, 1).Value = ""
            Sheet_IT_7B_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_7B_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_7B_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_7B_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_7B_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_7B_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_7B_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_7B_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_7B_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_7B_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_7B_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_7B_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_7B_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_7B_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_7B_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_7B_G3.Cells(2, 2).Value = frmIT_7B_G3.cmb1.Text
            Sheet_IT_7B_G3.Cells(3, 2).Value = frmIT_7B_G3.cmb2.Text
            Sheet_IT_7B_G3.Cells(4, 2).Value = frmIT_7B_G3.cmb3.Text
            Sheet_IT_7B_G3.Cells(5, 2).Value = frmIT_7B_G3.cmb4.Text
            Sheet_IT_7B_G3.Cells(6, 2).Value = frmIT_7B_G3.cmb5.Text
            Sheet_IT_7B_G3.Cells(7, 2).Value = frmIT_7B_G3.cmb6.Text
            Sheet_IT_7B_G3.Cells(8, 2).Value = frmIT_7B_G3.cmb7.Text
            Sheet_IT_7B_G3.Cells(9, 2).Value = frmIT_7B_G3.cmb8.Text
            Sheet_IT_7B_G3.Cells(10, 2).Value = frmIT_7B_G3.cmb9.Text
            Sheet_IT_7B_G3.Cells(11, 2).Value = frmIT_7B_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_7B_G3.Cells(2, 3).Value = frmIT_7B_G3.cmb21.Text
            Sheet_IT_7B_G3.Cells(3, 3).Value = frmIT_7B_G3.cmb22.Text
            Sheet_IT_7B_G3.Cells(4, 3).Value = frmIT_7B_G3.cmb23.Text
            Sheet_IT_7B_G3.Cells(5, 3).Value = frmIT_7B_G3.cmb24.Text
            Sheet_IT_7B_G3.Cells(6, 3).Value = frmIT_7B_G3.cmb25.Text
            Sheet_IT_7B_G3.Cells(7, 3).Value = frmIT_7B_G3.cmb26.Text
            Sheet_IT_7B_G3.Cells(8, 3).Value = frmIT_7B_G3.cmb27.Text
            Sheet_IT_7B_G3.Cells(9, 3).Value = frmIT_7B_G3.cmb28.Text
            Sheet_IT_7B_G3.Cells(10, 3).Value = frmIT_7B_G3.cmb29.Text
            Sheet_IT_7B_G3.Cells(11, 3).Value = frmIT_7B_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_7B_G3.Cells(2, 4).Value = frmIT_7B_G3.cmb41.Text
            Sheet_IT_7B_G3.Cells(3, 4).Value = frmIT_7B_G3.cmb42.Text
            Sheet_IT_7B_G3.Cells(4, 4).Value = frmIT_7B_G3.cmb43.Text
            Sheet_IT_7B_G3.Cells(5, 4).Value = frmIT_7B_G3.cmb44.Text
            Sheet_IT_7B_G3.Cells(6, 4).Value = frmIT_7B_G3.cmb45.Text
            Sheet_IT_7B_G3.Cells(7, 4).Value = frmIT_7B_G3.cmb46.Text
            Sheet_IT_7B_G3.Cells(8, 4).Value = frmIT_7B_G3.cmb47.Text
            Sheet_IT_7B_G3.Cells(9, 4).Value = frmIT_7B_G3.cmb48.Text
            Sheet_IT_7B_G3.Cells(10, 4).Value = frmIT_7B_G3.cmb49.Text
            Sheet_IT_7B_G3.Cells(11, 4).Value = frmIT_7B_G3.cmb50.Text

            ' Thursday

            Sheet_IT_7B_G3.Cells(2, 5).Value = frmIT_7B_G3.cmb61.Text
            Sheet_IT_7B_G3.Cells(3, 5).Value = frmIT_7B_G3.cmb62.Text
            Sheet_IT_7B_G3.Cells(4, 5).Value = frmIT_7B_G3.cmb63.Text
            Sheet_IT_7B_G3.Cells(5, 5).Value = frmIT_7B_G3.cmb64.Text
            Sheet_IT_7B_G3.Cells(6, 5).Value = frmIT_7B_G3.cmb65.Text
            Sheet_IT_7B_G3.Cells(7, 5).Value = frmIT_7B_G3.cmb66.Text
            Sheet_IT_7B_G3.Cells(8, 5).Value = frmIT_7B_G3.cmb67.Text
            Sheet_IT_7B_G3.Cells(9, 5).Value = frmIT_7B_G3.cmb68.Text
            Sheet_IT_7B_G3.Cells(10, 5).Value = frmIT_7B_G3.cmb69.Text
            Sheet_IT_7B_G3.Cells(11, 5).Value = frmIT_7B_G3.cmb70.Text

            ' Friday

            Sheet_IT_7B_G3.Cells(2, 6).Value = frmIT_7B_G3.cmb81.Text
            Sheet_IT_7B_G3.Cells(3, 6).Value = frmIT_7B_G3.cmb82.Text
            Sheet_IT_7B_G3.Cells(4, 6).Value = frmIT_7B_G3.cmb83.Text
            Sheet_IT_7B_G3.Cells(5, 6).Value = frmIT_7B_G3.cmb84.Text
            Sheet_IT_7B_G3.Cells(6, 6).Value = frmIT_7B_G3.cmb85.Text
            Sheet_IT_7B_G3.Cells(7, 6).Value = frmIT_7B_G3.cmb86.Text
            Sheet_IT_7B_G3.Cells(8, 6).Value = frmIT_7B_G3.cmb87.Text
            Sheet_IT_7B_G3.Cells(9, 6).Value = frmIT_7B_G3.cmb88.Text
            Sheet_IT_7B_G3.Cells(10, 6).Value = frmIT_7B_G3.cmb89.Text
            Sheet_IT_7B_G3.Cells(11, 6).Value = frmIT_7B_G3.cmb90.Text

            Sheet_IT_3A_G1 = Nothing
            Sheet_IT_3A_G2 = Nothing
            Sheet_IT_3A_G3 = Nothing
            Sheet_IT_3B_G1 = Nothing
            Sheet_IT_3B_G2 = Nothing
            Sheet_IT_3B_G3 = Nothing

            Sheet_IT_5A_G1 = Nothing
            Sheet_IT_5A_G2 = Nothing
            Sheet_IT_5A_G3 = Nothing
            Sheet_IT_5B_G1 = Nothing
            Sheet_IT_5B_G2 = Nothing
            Sheet_IT_5B_G3 = Nothing

            Sheet_IT_7A_G1 = Nothing
            Sheet_IT_7A_G2 = Nothing
            Sheet_IT_7A_G3 = Nothing
            Sheet_IT_7B_G1 = Nothing
            Sheet_IT_7B_G2 = Nothing
            Sheet_IT_7B_G3 = Nothing

        Else
            Dim Sheet_IT_4A_G1 As Excel.Worksheet
            Dim Sheet_IT_4A_G2 As Excel.Worksheet
            Dim Sheet_IT_4A_G3 As Excel.Worksheet
            Dim Sheet_IT_4B_G1 As Excel.Worksheet
            Dim Sheet_IT_4B_G2 As Excel.Worksheet
            Dim Sheet_IT_4B_G3 As Excel.Worksheet

            Dim Sheet_IT_6A_G1 As Excel.Worksheet
            Dim Sheet_IT_6A_G2 As Excel.Worksheet
            Dim Sheet_IT_6A_G3 As Excel.Worksheet
            Dim Sheet_IT_6B_G1 As Excel.Worksheet
            Dim Sheet_IT_6B_G2 As Excel.Worksheet
            Dim Sheet_IT_6B_G3 As Excel.Worksheet

            Dim Sheet_IT_8A_G1 As Excel.Worksheet
            Dim Sheet_IT_8A_G2 As Excel.Worksheet
            Dim Sheet_IT_8A_G3 As Excel.Worksheet
            Dim Sheet_IT_8B_G1 As Excel.Worksheet
            Dim Sheet_IT_8B_G2 As Excel.Worksheet
            Dim Sheet_IT_8B_G3 As Excel.Worksheet

            ' Get new workbooks.

            oWB = oXL.Workbooks.Add
            Sheet_IT_4A_G1 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_4A_G2 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_4A_G3 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_4B_G1 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_4B_G2 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_4B_G3 = oWB.ActiveSheet

            Sheet_IT_6A_G1 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_6A_G2 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_6A_G3 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_6B_G1 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_6B_G2 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_6B_G3 = oWB.ActiveSheet

            Sheet_IT_8A_G1 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_8A_G2 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_8A_G3 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_8B_G1 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_8B_G2 = oWB.ActiveSheet
            oWB = oXL.Workbooks.Add
            Sheet_IT_8B_G3 = oWB.ActiveSheet

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4A_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4A_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4A_G1.Cells(1, 1).Value = ""
            Sheet_IT_4A_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4A_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4A_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4A_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4A_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4A_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4A_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4A_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4A_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4A_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4A_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_4A_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4A_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4A_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4A_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4A_G1.Cells(2, 2).Value = frmIT_4A_G1.cmb1.Text
            Sheet_IT_4A_G1.Cells(3, 2).Value = frmIT_4A_G1.cmb2.Text
            Sheet_IT_4A_G1.Cells(4, 2).Value = frmIT_4A_G1.cmb3.Text
            Sheet_IT_4A_G1.Cells(5, 2).Value = frmIT_4A_G1.cmb4.Text
            Sheet_IT_4A_G1.Cells(6, 2).Value = frmIT_4A_G1.cmb5.Text
            Sheet_IT_4A_G1.Cells(7, 2).Value = frmIT_4A_G1.cmb6.Text
            Sheet_IT_4A_G1.Cells(8, 2).Value = frmIT_4A_G1.cmb7.Text
            Sheet_IT_4A_G1.Cells(9, 2).Value = frmIT_4A_G1.cmb8.Text
            Sheet_IT_4A_G1.Cells(10, 2).Value = frmIT_4A_G1.cmb9.Text
            Sheet_IT_4A_G1.Cells(11, 2).Value = frmIT_4A_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_4A_G1.Cells(2, 3).Value = frmIT_4A_G1.cmb21.Text
            Sheet_IT_4A_G1.Cells(3, 3).Value = frmIT_4A_G1.cmb22.Text
            Sheet_IT_4A_G1.Cells(4, 3).Value = frmIT_4A_G1.cmb23.Text
            Sheet_IT_4A_G1.Cells(5, 3).Value = frmIT_4A_G1.cmb24.Text
            Sheet_IT_4A_G1.Cells(6, 3).Value = frmIT_4A_G1.cmb25.Text
            Sheet_IT_4A_G1.Cells(7, 3).Value = frmIT_4A_G1.cmb26.Text
            Sheet_IT_4A_G1.Cells(8, 3).Value = frmIT_4A_G1.cmb27.Text
            Sheet_IT_4A_G1.Cells(9, 3).Value = frmIT_4A_G1.cmb28.Text
            Sheet_IT_4A_G1.Cells(10, 3).Value = frmIT_4A_G1.cmb29.Text
            Sheet_IT_4A_G1.Cells(11, 3).Value = frmIT_4A_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_4A_G1.Cells(2, 4).Value = frmIT_4A_G1.cmb41.Text
            Sheet_IT_4A_G1.Cells(3, 4).Value = frmIT_4A_G1.cmb42.Text
            Sheet_IT_4A_G1.Cells(4, 4).Value = frmIT_4A_G1.cmb43.Text
            Sheet_IT_4A_G1.Cells(5, 4).Value = frmIT_4A_G1.cmb44.Text
            Sheet_IT_4A_G1.Cells(6, 4).Value = frmIT_4A_G1.cmb45.Text
            Sheet_IT_4A_G1.Cells(7, 4).Value = frmIT_4A_G1.cmb46.Text
            Sheet_IT_4A_G1.Cells(8, 4).Value = frmIT_4A_G1.cmb47.Text
            Sheet_IT_4A_G1.Cells(9, 4).Value = frmIT_4A_G1.cmb48.Text
            Sheet_IT_4A_G1.Cells(10, 4).Value = frmIT_4A_G1.cmb49.Text
            Sheet_IT_4A_G1.Cells(11, 4).Value = frmIT_4A_G1.cmb50.Text

            ' Thursday

            Sheet_IT_4A_G1.Cells(2, 5).Value = frmIT_4A_G1.cmb61.Text
            Sheet_IT_4A_G1.Cells(3, 5).Value = frmIT_4A_G1.cmb62.Text
            Sheet_IT_4A_G1.Cells(4, 5).Value = frmIT_4A_G1.cmb63.Text
            Sheet_IT_4A_G1.Cells(5, 5).Value = frmIT_4A_G1.cmb64.Text
            Sheet_IT_4A_G1.Cells(6, 5).Value = frmIT_4A_G1.cmb65.Text
            Sheet_IT_4A_G1.Cells(7, 5).Value = frmIT_4A_G1.cmb66.Text
            Sheet_IT_4A_G1.Cells(8, 5).Value = frmIT_4A_G1.cmb67.Text
            Sheet_IT_4A_G1.Cells(9, 5).Value = frmIT_4A_G1.cmb68.Text
            Sheet_IT_4A_G1.Cells(10, 5).Value = frmIT_4A_G1.cmb69.Text
            Sheet_IT_4A_G1.Cells(11, 5).Value = frmIT_4A_G1.cmb70.Text

            ' Friday

            Sheet_IT_4A_G1.Cells(2, 6).Value = frmIT_4A_G1.cmb81.Text
            Sheet_IT_4A_G1.Cells(3, 6).Value = frmIT_4A_G1.cmb82.Text
            Sheet_IT_4A_G1.Cells(4, 6).Value = frmIT_4A_G1.cmb83.Text
            Sheet_IT_4A_G1.Cells(5, 6).Value = frmIT_4A_G1.cmb84.Text
            Sheet_IT_4A_G1.Cells(6, 6).Value = frmIT_4A_G1.cmb85.Text
            Sheet_IT_4A_G1.Cells(7, 6).Value = frmIT_4A_G1.cmb86.Text
            Sheet_IT_4A_G1.Cells(8, 6).Value = frmIT_4A_G1.cmb87.Text
            Sheet_IT_4A_G1.Cells(9, 6).Value = frmIT_4A_G1.cmb88.Text
            Sheet_IT_4A_G1.Cells(10, 6).Value = frmIT_4A_G1.cmb89.Text
            Sheet_IT_4A_G1.Cells(11, 6).Value = frmIT_4A_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4A_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4A_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4A_G2.Cells(1, 1).Value = ""
            Sheet_IT_4A_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4A_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4A_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4A_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4A_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4A_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4A_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4A_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4A_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4A_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4A_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_4A_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4A_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4A_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4A_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4A_G2.Cells(2, 2).Value = frmIT_4A_G2.cmb1.Text
            Sheet_IT_4A_G2.Cells(3, 2).Value = frmIT_4A_G2.cmb2.Text
            Sheet_IT_4A_G2.Cells(4, 2).Value = frmIT_4A_G2.cmb3.Text
            Sheet_IT_4A_G2.Cells(5, 2).Value = frmIT_4A_G2.cmb4.Text
            Sheet_IT_4A_G2.Cells(6, 2).Value = frmIT_4A_G2.cmb5.Text
            Sheet_IT_4A_G2.Cells(7, 2).Value = frmIT_4A_G2.cmb6.Text
            Sheet_IT_4A_G2.Cells(8, 2).Value = frmIT_4A_G2.cmb7.Text
            Sheet_IT_4A_G2.Cells(9, 2).Value = frmIT_4A_G2.cmb8.Text
            Sheet_IT_4A_G2.Cells(10, 2).Value = frmIT_4A_G2.cmb9.Text
            Sheet_IT_4A_G2.Cells(11, 2).Value = frmIT_4A_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_4A_G2.Cells(2, 3).Value = frmIT_4A_G2.cmb21.Text
            Sheet_IT_4A_G2.Cells(3, 3).Value = frmIT_4A_G2.cmb22.Text
            Sheet_IT_4A_G2.Cells(4, 3).Value = frmIT_4A_G2.cmb23.Text
            Sheet_IT_4A_G2.Cells(5, 3).Value = frmIT_4A_G2.cmb24.Text
            Sheet_IT_4A_G2.Cells(6, 3).Value = frmIT_4A_G2.cmb25.Text
            Sheet_IT_4A_G2.Cells(7, 3).Value = frmIT_4A_G2.cmb26.Text
            Sheet_IT_4A_G2.Cells(8, 3).Value = frmIT_4A_G2.cmb27.Text
            Sheet_IT_4A_G2.Cells(9, 3).Value = frmIT_4A_G2.cmb28.Text
            Sheet_IT_4A_G2.Cells(10, 3).Value = frmIT_4A_G2.cmb29.Text
            Sheet_IT_4A_G2.Cells(11, 3).Value = frmIT_4A_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_4A_G2.Cells(2, 4).Value = frmIT_4A_G2.cmb41.Text
            Sheet_IT_4A_G2.Cells(3, 4).Value = frmIT_4A_G2.cmb42.Text
            Sheet_IT_4A_G2.Cells(4, 4).Value = frmIT_4A_G2.cmb43.Text
            Sheet_IT_4A_G2.Cells(5, 4).Value = frmIT_4A_G2.cmb44.Text
            Sheet_IT_4A_G2.Cells(6, 4).Value = frmIT_4A_G2.cmb45.Text
            Sheet_IT_4A_G2.Cells(7, 4).Value = frmIT_4A_G2.cmb46.Text
            Sheet_IT_4A_G2.Cells(8, 4).Value = frmIT_4A_G2.cmb47.Text
            Sheet_IT_4A_G2.Cells(9, 4).Value = frmIT_4A_G2.cmb48.Text
            Sheet_IT_4A_G2.Cells(10, 4).Value = frmIT_4A_G2.cmb49.Text
            Sheet_IT_4A_G2.Cells(11, 4).Value = frmIT_4A_G2.cmb50.Text

            ' Thursday

            Sheet_IT_4A_G2.Cells(2, 5).Value = frmIT_4A_G2.cmb61.Text
            Sheet_IT_4A_G2.Cells(3, 5).Value = frmIT_4A_G2.cmb62.Text
            Sheet_IT_4A_G2.Cells(4, 5).Value = frmIT_4A_G2.cmb63.Text
            Sheet_IT_4A_G2.Cells(5, 5).Value = frmIT_4A_G2.cmb64.Text
            Sheet_IT_4A_G2.Cells(6, 5).Value = frmIT_4A_G2.cmb65.Text
            Sheet_IT_4A_G2.Cells(7, 5).Value = frmIT_4A_G2.cmb66.Text
            Sheet_IT_4A_G2.Cells(8, 5).Value = frmIT_4A_G2.cmb67.Text
            Sheet_IT_4A_G2.Cells(9, 5).Value = frmIT_4A_G2.cmb68.Text
            Sheet_IT_4A_G2.Cells(10, 5).Value = frmIT_4A_G2.cmb69.Text
            Sheet_IT_4A_G2.Cells(11, 5).Value = frmIT_4A_G2.cmb70.Text

            ' Friday

            Sheet_IT_4A_G2.Cells(2, 6).Value = frmIT_4A_G2.cmb81.Text
            Sheet_IT_4A_G2.Cells(3, 6).Value = frmIT_4A_G2.cmb82.Text
            Sheet_IT_4A_G2.Cells(4, 6).Value = frmIT_4A_G2.cmb83.Text
            Sheet_IT_4A_G2.Cells(5, 6).Value = frmIT_4A_G2.cmb84.Text
            Sheet_IT_4A_G2.Cells(6, 6).Value = frmIT_4A_G2.cmb85.Text
            Sheet_IT_4A_G2.Cells(7, 6).Value = frmIT_4A_G2.cmb86.Text
            Sheet_IT_4A_G2.Cells(8, 6).Value = frmIT_4A_G2.cmb87.Text
            Sheet_IT_4A_G2.Cells(9, 6).Value = frmIT_4A_G2.cmb88.Text
            Sheet_IT_4A_G2.Cells(10, 6).Value = frmIT_4A_G2.cmb89.Text
            Sheet_IT_4A_G2.Cells(11, 6).Value = frmIT_4A_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4A_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4A_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4A_G3.Cells(1, 1).Value = ""
            Sheet_IT_4A_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4A_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4A_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4A_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4A_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4A_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4A_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4A_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4A_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4A_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4A_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_4A_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4A_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4A_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4A_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4A_G3.Cells(2, 2).Value = frmIT_4A_G3.cmb1.Text
            Sheet_IT_4A_G3.Cells(3, 2).Value = frmIT_4A_G3.cmb2.Text
            Sheet_IT_4A_G3.Cells(4, 2).Value = frmIT_4A_G3.cmb3.Text
            Sheet_IT_4A_G3.Cells(5, 2).Value = frmIT_4A_G3.cmb4.Text
            Sheet_IT_4A_G3.Cells(6, 2).Value = frmIT_4A_G3.cmb5.Text
            Sheet_IT_4A_G3.Cells(7, 2).Value = frmIT_4A_G3.cmb6.Text
            Sheet_IT_4A_G3.Cells(8, 2).Value = frmIT_4A_G3.cmb7.Text
            Sheet_IT_4A_G3.Cells(9, 2).Value = frmIT_4A_G3.cmb8.Text
            Sheet_IT_4A_G3.Cells(10, 2).Value = frmIT_4A_G3.cmb9.Text
            Sheet_IT_4A_G3.Cells(11, 2).Value = frmIT_4A_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_4A_G3.Cells(2, 3).Value = frmIT_4A_G3.cmb21.Text
            Sheet_IT_4A_G3.Cells(3, 3).Value = frmIT_4A_G3.cmb22.Text
            Sheet_IT_4A_G3.Cells(4, 3).Value = frmIT_4A_G3.cmb23.Text
            Sheet_IT_4A_G3.Cells(5, 3).Value = frmIT_4A_G3.cmb24.Text
            Sheet_IT_4A_G3.Cells(6, 3).Value = frmIT_4A_G3.cmb25.Text
            Sheet_IT_4A_G3.Cells(7, 3).Value = frmIT_4A_G3.cmb26.Text
            Sheet_IT_4A_G3.Cells(8, 3).Value = frmIT_4A_G3.cmb27.Text
            Sheet_IT_4A_G3.Cells(9, 3).Value = frmIT_4A_G3.cmb28.Text
            Sheet_IT_4A_G3.Cells(10, 3).Value = frmIT_4A_G3.cmb29.Text
            Sheet_IT_4A_G3.Cells(11, 3).Value = frmIT_4A_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_4A_G3.Cells(2, 4).Value = frmIT_4A_G3.cmb41.Text
            Sheet_IT_4A_G3.Cells(3, 4).Value = frmIT_4A_G3.cmb42.Text
            Sheet_IT_4A_G3.Cells(4, 4).Value = frmIT_4A_G3.cmb43.Text
            Sheet_IT_4A_G3.Cells(5, 4).Value = frmIT_4A_G3.cmb44.Text
            Sheet_IT_4A_G3.Cells(6, 4).Value = frmIT_4A_G3.cmb45.Text
            Sheet_IT_4A_G3.Cells(7, 4).Value = frmIT_4A_G3.cmb46.Text
            Sheet_IT_4A_G3.Cells(8, 4).Value = frmIT_4A_G3.cmb47.Text
            Sheet_IT_4A_G3.Cells(9, 4).Value = frmIT_4A_G3.cmb48.Text
            Sheet_IT_4A_G3.Cells(10, 4).Value = frmIT_4A_G3.cmb49.Text
            Sheet_IT_4A_G3.Cells(11, 4).Value = frmIT_4A_G3.cmb50.Text

            ' Thursday

            Sheet_IT_4A_G3.Cells(2, 5).Value = frmIT_4A_G3.cmb61.Text
            Sheet_IT_4A_G3.Cells(3, 5).Value = frmIT_4A_G3.cmb62.Text
            Sheet_IT_4A_G3.Cells(4, 5).Value = frmIT_4A_G3.cmb63.Text
            Sheet_IT_4A_G3.Cells(5, 5).Value = frmIT_4A_G3.cmb64.Text
            Sheet_IT_4A_G3.Cells(6, 5).Value = frmIT_4A_G3.cmb65.Text
            Sheet_IT_4A_G3.Cells(7, 5).Value = frmIT_4A_G3.cmb66.Text
            Sheet_IT_4A_G3.Cells(8, 5).Value = frmIT_4A_G3.cmb67.Text
            Sheet_IT_4A_G3.Cells(9, 5).Value = frmIT_4A_G3.cmb68.Text
            Sheet_IT_4A_G3.Cells(10, 5).Value = frmIT_4A_G3.cmb69.Text
            Sheet_IT_4A_G3.Cells(11, 5).Value = frmIT_4A_G3.cmb70.Text

            ' Friday

            Sheet_IT_4A_G3.Cells(2, 6).Value = frmIT_4A_G3.cmb81.Text
            Sheet_IT_4A_G3.Cells(3, 6).Value = frmIT_4A_G3.cmb82.Text
            Sheet_IT_4A_G3.Cells(4, 6).Value = frmIT_4A_G3.cmb83.Text
            Sheet_IT_4A_G3.Cells(5, 6).Value = frmIT_4A_G3.cmb84.Text
            Sheet_IT_4A_G3.Cells(6, 6).Value = frmIT_4A_G3.cmb85.Text
            Sheet_IT_4A_G3.Cells(7, 6).Value = frmIT_4A_G3.cmb86.Text
            Sheet_IT_4A_G3.Cells(8, 6).Value = frmIT_4A_G3.cmb87.Text
            Sheet_IT_4A_G3.Cells(9, 6).Value = frmIT_4A_G3.cmb88.Text
            Sheet_IT_4A_G3.Cells(10, 6).Value = frmIT_4A_G3.cmb89.Text
            Sheet_IT_4A_G3.Cells(11, 6).Value = frmIT_4A_G3.cmb90.Text
            ''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''
            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4B_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4B_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4B_G1.Cells(1, 1).Value = ""
            Sheet_IT_4B_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4B_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4B_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4B_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4B_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4B_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4B_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4B_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4B_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4B_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4B_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_4B_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4B_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4B_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4B_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4B_G1.Cells(2, 2).Value = frmIT_4B_G1.cmb1.Text
            Sheet_IT_4B_G1.Cells(3, 2).Value = frmIT_4B_G1.cmb2.Text
            Sheet_IT_4B_G1.Cells(4, 2).Value = frmIT_4B_G1.cmb3.Text
            Sheet_IT_4B_G1.Cells(5, 2).Value = frmIT_4B_G1.cmb4.Text
            Sheet_IT_4B_G1.Cells(6, 2).Value = frmIT_4B_G1.cmb5.Text
            Sheet_IT_4B_G1.Cells(7, 2).Value = frmIT_4B_G1.cmb6.Text
            Sheet_IT_4B_G1.Cells(8, 2).Value = frmIT_4B_G1.cmb7.Text
            Sheet_IT_4B_G1.Cells(9, 2).Value = frmIT_4B_G1.cmb8.Text
            Sheet_IT_4B_G1.Cells(10, 2).Value = frmIT_4B_G1.cmb9.Text
            Sheet_IT_4B_G1.Cells(11, 2).Value = frmIT_4B_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_4B_G1.Cells(2, 3).Value = frmIT_4B_G1.cmb21.Text
            Sheet_IT_4B_G1.Cells(3, 3).Value = frmIT_4B_G1.cmb22.Text
            Sheet_IT_4B_G1.Cells(4, 3).Value = frmIT_4B_G1.cmb23.Text
            Sheet_IT_4B_G1.Cells(5, 3).Value = frmIT_4B_G1.cmb24.Text
            Sheet_IT_4B_G1.Cells(6, 3).Value = frmIT_4B_G1.cmb25.Text
            Sheet_IT_4B_G1.Cells(7, 3).Value = frmIT_4B_G1.cmb26.Text
            Sheet_IT_4B_G1.Cells(8, 3).Value = frmIT_4B_G1.cmb27.Text
            Sheet_IT_4B_G1.Cells(9, 3).Value = frmIT_4B_G1.cmb28.Text
            Sheet_IT_4B_G1.Cells(10, 3).Value = frmIT_4B_G1.cmb29.Text
            Sheet_IT_4B_G1.Cells(11, 3).Value = frmIT_4B_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_4B_G1.Cells(2, 4).Value = frmIT_4B_G1.cmb41.Text
            Sheet_IT_4B_G1.Cells(3, 4).Value = frmIT_4B_G1.cmb42.Text
            Sheet_IT_4B_G1.Cells(4, 4).Value = frmIT_4B_G1.cmb43.Text
            Sheet_IT_4B_G1.Cells(5, 4).Value = frmIT_4B_G1.cmb44.Text
            Sheet_IT_4B_G1.Cells(6, 4).Value = frmIT_4B_G1.cmb45.Text
            Sheet_IT_4B_G1.Cells(7, 4).Value = frmIT_4B_G1.cmb46.Text
            Sheet_IT_4B_G1.Cells(8, 4).Value = frmIT_4B_G1.cmb47.Text
            Sheet_IT_4B_G1.Cells(9, 4).Value = frmIT_4B_G1.cmb48.Text
            Sheet_IT_4B_G1.Cells(10, 4).Value = frmIT_4B_G1.cmb49.Text
            Sheet_IT_4B_G1.Cells(11, 4).Value = frmIT_4B_G1.cmb50.Text

            ' Thursday

            Sheet_IT_4B_G1.Cells(2, 5).Value = frmIT_4B_G1.cmb61.Text
            Sheet_IT_4B_G1.Cells(3, 5).Value = frmIT_4B_G1.cmb62.Text
            Sheet_IT_4B_G1.Cells(4, 5).Value = frmIT_4B_G1.cmb63.Text
            Sheet_IT_4B_G1.Cells(5, 5).Value = frmIT_4B_G1.cmb64.Text
            Sheet_IT_4B_G1.Cells(6, 5).Value = frmIT_4B_G1.cmb65.Text
            Sheet_IT_4B_G1.Cells(7, 5).Value = frmIT_4B_G1.cmb66.Text
            Sheet_IT_4B_G1.Cells(8, 5).Value = frmIT_4B_G1.cmb67.Text
            Sheet_IT_4B_G1.Cells(9, 5).Value = frmIT_4B_G1.cmb68.Text
            Sheet_IT_4B_G1.Cells(10, 5).Value = frmIT_4B_G1.cmb69.Text
            Sheet_IT_4B_G1.Cells(11, 5).Value = frmIT_4B_G1.cmb70.Text

            ' Friday

            Sheet_IT_4B_G1.Cells(2, 6).Value = frmIT_4B_G1.cmb81.Text
            Sheet_IT_4B_G1.Cells(3, 6).Value = frmIT_4B_G1.cmb82.Text
            Sheet_IT_4B_G1.Cells(4, 6).Value = frmIT_4B_G1.cmb83.Text
            Sheet_IT_4B_G1.Cells(5, 6).Value = frmIT_4B_G1.cmb84.Text
            Sheet_IT_4B_G1.Cells(6, 6).Value = frmIT_4B_G1.cmb85.Text
            Sheet_IT_4B_G1.Cells(7, 6).Value = frmIT_4B_G1.cmb86.Text
            Sheet_IT_4B_G1.Cells(8, 6).Value = frmIT_4B_G1.cmb87.Text
            Sheet_IT_4B_G1.Cells(9, 6).Value = frmIT_4B_G1.cmb88.Text
            Sheet_IT_4B_G1.Cells(10, 6).Value = frmIT_4B_G1.cmb89.Text
            Sheet_IT_4B_G1.Cells(11, 6).Value = frmIT_4B_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4B_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4B_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4B_G2.Cells(1, 1).Value = ""
            Sheet_IT_4B_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4B_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4B_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4B_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4B_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4B_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4B_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4B_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4B_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4B_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4B_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_4B_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4B_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4B_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4B_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4B_G2.Cells(2, 2).Value = frmIT_4B_G2.cmb1.Text
            Sheet_IT_4B_G2.Cells(3, 2).Value = frmIT_4B_G2.cmb2.Text
            Sheet_IT_4B_G2.Cells(4, 2).Value = frmIT_4B_G2.cmb3.Text
            Sheet_IT_4B_G2.Cells(5, 2).Value = frmIT_4B_G2.cmb4.Text
            Sheet_IT_4B_G2.Cells(6, 2).Value = frmIT_4B_G2.cmb5.Text
            Sheet_IT_4B_G2.Cells(7, 2).Value = frmIT_4B_G2.cmb6.Text
            Sheet_IT_4B_G2.Cells(8, 2).Value = frmIT_4B_G2.cmb7.Text
            Sheet_IT_4B_G2.Cells(9, 2).Value = frmIT_4B_G2.cmb8.Text
            Sheet_IT_4B_G2.Cells(10, 2).Value = frmIT_4B_G2.cmb9.Text
            Sheet_IT_4B_G2.Cells(11, 2).Value = frmIT_4B_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_4B_G2.Cells(2, 3).Value = frmIT_4B_G2.cmb21.Text
            Sheet_IT_4B_G2.Cells(3, 3).Value = frmIT_4B_G2.cmb22.Text
            Sheet_IT_4B_G2.Cells(4, 3).Value = frmIT_4B_G2.cmb23.Text
            Sheet_IT_4B_G2.Cells(5, 3).Value = frmIT_4B_G2.cmb24.Text
            Sheet_IT_4B_G2.Cells(6, 3).Value = frmIT_4B_G2.cmb25.Text
            Sheet_IT_4B_G2.Cells(7, 3).Value = frmIT_4B_G2.cmb26.Text
            Sheet_IT_4B_G2.Cells(8, 3).Value = frmIT_4B_G2.cmb27.Text
            Sheet_IT_4B_G2.Cells(9, 3).Value = frmIT_4B_G2.cmb28.Text
            Sheet_IT_4B_G2.Cells(10, 3).Value = frmIT_4B_G2.cmb29.Text
            Sheet_IT_4B_G2.Cells(11, 3).Value = frmIT_4B_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_4B_G2.Cells(2, 4).Value = frmIT_4B_G2.cmb41.Text
            Sheet_IT_4B_G2.Cells(3, 4).Value = frmIT_4B_G2.cmb42.Text
            Sheet_IT_4B_G2.Cells(4, 4).Value = frmIT_4B_G2.cmb43.Text
            Sheet_IT_4B_G2.Cells(5, 4).Value = frmIT_4B_G2.cmb44.Text
            Sheet_IT_4B_G2.Cells(6, 4).Value = frmIT_4B_G2.cmb45.Text
            Sheet_IT_4B_G2.Cells(7, 4).Value = frmIT_4B_G2.cmb46.Text
            Sheet_IT_4B_G2.Cells(8, 4).Value = frmIT_4B_G2.cmb47.Text
            Sheet_IT_4B_G2.Cells(9, 4).Value = frmIT_4B_G2.cmb48.Text
            Sheet_IT_4B_G2.Cells(10, 4).Value = frmIT_4B_G2.cmb49.Text
            Sheet_IT_4B_G2.Cells(11, 4).Value = frmIT_4B_G2.cmb50.Text

            ' Thursday

            Sheet_IT_4B_G2.Cells(2, 5).Value = frmIT_4B_G2.cmb61.Text
            Sheet_IT_4B_G2.Cells(3, 5).Value = frmIT_4B_G2.cmb62.Text
            Sheet_IT_4B_G2.Cells(4, 5).Value = frmIT_4B_G2.cmb63.Text
            Sheet_IT_4B_G2.Cells(5, 5).Value = frmIT_4B_G2.cmb64.Text
            Sheet_IT_4B_G2.Cells(6, 5).Value = frmIT_4B_G2.cmb65.Text
            Sheet_IT_4B_G2.Cells(7, 5).Value = frmIT_4B_G2.cmb66.Text
            Sheet_IT_4B_G2.Cells(8, 5).Value = frmIT_4B_G2.cmb67.Text
            Sheet_IT_4B_G2.Cells(9, 5).Value = frmIT_4B_G2.cmb68.Text
            Sheet_IT_4B_G2.Cells(10, 5).Value = frmIT_4B_G2.cmb69.Text
            Sheet_IT_4B_G2.Cells(11, 5).Value = frmIT_4B_G2.cmb70.Text

            ' Friday

            Sheet_IT_4B_G2.Cells(2, 6).Value = frmIT_4B_G2.cmb81.Text
            Sheet_IT_4B_G2.Cells(3, 6).Value = frmIT_4B_G2.cmb82.Text
            Sheet_IT_4B_G2.Cells(4, 6).Value = frmIT_4B_G2.cmb83.Text
            Sheet_IT_4B_G2.Cells(5, 6).Value = frmIT_4B_G2.cmb84.Text
            Sheet_IT_4B_G2.Cells(6, 6).Value = frmIT_4B_G2.cmb85.Text
            Sheet_IT_4B_G2.Cells(7, 6).Value = frmIT_4B_G2.cmb86.Text
            Sheet_IT_4B_G2.Cells(8, 6).Value = frmIT_4B_G2.cmb87.Text
            Sheet_IT_4B_G2.Cells(9, 6).Value = frmIT_4B_G2.cmb88.Text
            Sheet_IT_4B_G2.Cells(10, 6).Value = frmIT_4B_G2.cmb89.Text
            Sheet_IT_4B_G2.Cells(11, 6).Value = frmIT_4B_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4B_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4B_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4B_G3.Cells(1, 1).Value = ""
            Sheet_IT_4B_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4B_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4B_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4B_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4B_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4B_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4B_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4B_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4B_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4B_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4B_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_4B_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4B_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4B_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4B_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4B_G3.Cells(2, 2).Value = frmIT_4B_G3.cmb1.Text
            Sheet_IT_4B_G3.Cells(3, 2).Value = frmIT_4B_G3.cmb2.Text
            Sheet_IT_4B_G3.Cells(4, 2).Value = frmIT_4B_G3.cmb3.Text
            Sheet_IT_4B_G3.Cells(5, 2).Value = frmIT_4B_G3.cmb4.Text
            Sheet_IT_4B_G3.Cells(6, 2).Value = frmIT_4B_G3.cmb5.Text
            Sheet_IT_4B_G3.Cells(7, 2).Value = frmIT_4B_G3.cmb6.Text
            Sheet_IT_4B_G3.Cells(8, 2).Value = frmIT_4B_G3.cmb7.Text
            Sheet_IT_4B_G3.Cells(9, 2).Value = frmIT_4B_G3.cmb8.Text
            Sheet_IT_4B_G3.Cells(10, 2).Value = frmIT_4B_G3.cmb9.Text
            Sheet_IT_4B_G3.Cells(11, 2).Value = frmIT_4B_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_4B_G3.Cells(2, 3).Value = frmIT_4B_G3.cmb21.Text
            Sheet_IT_4B_G3.Cells(3, 3).Value = frmIT_4B_G3.cmb22.Text
            Sheet_IT_4B_G3.Cells(4, 3).Value = frmIT_4B_G3.cmb23.Text
            Sheet_IT_4B_G3.Cells(5, 3).Value = frmIT_4B_G3.cmb24.Text
            Sheet_IT_4B_G3.Cells(6, 3).Value = frmIT_4B_G3.cmb25.Text
            Sheet_IT_4B_G3.Cells(7, 3).Value = frmIT_4B_G3.cmb26.Text
            Sheet_IT_4B_G3.Cells(8, 3).Value = frmIT_4B_G3.cmb27.Text
            Sheet_IT_4B_G3.Cells(9, 3).Value = frmIT_4B_G3.cmb28.Text
            Sheet_IT_4B_G3.Cells(10, 3).Value = frmIT_4B_G3.cmb29.Text
            Sheet_IT_4B_G3.Cells(11, 3).Value = frmIT_4B_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_4B_G3.Cells(2, 4).Value = frmIT_4B_G3.cmb41.Text
            Sheet_IT_4B_G3.Cells(3, 4).Value = frmIT_4B_G3.cmb42.Text
            Sheet_IT_4B_G3.Cells(4, 4).Value = frmIT_4B_G3.cmb43.Text
            Sheet_IT_4B_G3.Cells(5, 4).Value = frmIT_4B_G3.cmb44.Text
            Sheet_IT_4B_G3.Cells(6, 4).Value = frmIT_4B_G3.cmb45.Text
            Sheet_IT_4B_G3.Cells(7, 4).Value = frmIT_4B_G3.cmb46.Text
            Sheet_IT_4B_G3.Cells(8, 4).Value = frmIT_4B_G3.cmb47.Text
            Sheet_IT_4B_G3.Cells(9, 4).Value = frmIT_4B_G3.cmb48.Text
            Sheet_IT_4B_G3.Cells(10, 4).Value = frmIT_4B_G3.cmb49.Text
            Sheet_IT_4B_G3.Cells(11, 4).Value = frmIT_4B_G3.cmb50.Text

            ' Thursday

            Sheet_IT_4B_G3.Cells(2, 5).Value = frmIT_4B_G3.cmb61.Text
            Sheet_IT_4B_G3.Cells(3, 5).Value = frmIT_4B_G3.cmb62.Text
            Sheet_IT_4B_G3.Cells(4, 5).Value = frmIT_4B_G3.cmb63.Text
            Sheet_IT_4B_G3.Cells(5, 5).Value = frmIT_4B_G3.cmb64.Text
            Sheet_IT_4B_G3.Cells(6, 5).Value = frmIT_4B_G3.cmb65.Text
            Sheet_IT_4B_G3.Cells(7, 5).Value = frmIT_4B_G3.cmb66.Text
            Sheet_IT_4B_G3.Cells(8, 5).Value = frmIT_4B_G3.cmb67.Text
            Sheet_IT_4B_G3.Cells(9, 5).Value = frmIT_4B_G3.cmb68.Text
            Sheet_IT_4B_G3.Cells(10, 5).Value = frmIT_4B_G3.cmb69.Text
            Sheet_IT_4B_G3.Cells(11, 5).Value = frmIT_4B_G3.cmb70.Text

            ' Friday

            Sheet_IT_4B_G3.Cells(2, 6).Value = frmIT_4B_G3.cmb81.Text
            Sheet_IT_4B_G3.Cells(3, 6).Value = frmIT_4B_G3.cmb82.Text
            Sheet_IT_4B_G3.Cells(4, 6).Value = frmIT_4B_G3.cmb83.Text
            Sheet_IT_4B_G3.Cells(5, 6).Value = frmIT_4B_G3.cmb84.Text
            Sheet_IT_4B_G3.Cells(6, 6).Value = frmIT_4B_G3.cmb85.Text
            Sheet_IT_4B_G3.Cells(7, 6).Value = frmIT_4B_G3.cmb86.Text
            Sheet_IT_4B_G3.Cells(8, 6).Value = frmIT_4B_G3.cmb87.Text
            Sheet_IT_4B_G3.Cells(9, 6).Value = frmIT_4B_G3.cmb88.Text
            Sheet_IT_4B_G3.Cells(10, 6).Value = frmIT_4B_G3.cmb89.Text
            Sheet_IT_4B_G3.Cells(11, 6).Value = frmIT_4B_G3.cmb90.Text
            '
            '
            '
            '
            '
            '
            '
            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_6A_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_6A_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_6A_G1.Cells(1, 1).Value = ""
            Sheet_IT_6A_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_6A_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_6A_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_6A_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_6A_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_6A_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_6A_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_6A_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_6A_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_6A_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_6A_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_6A_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_6A_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_6A_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_6A_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_6A_G1.Cells(2, 2).Value = frmIT_6A_G1.cmb1.Text
            Sheet_IT_6A_G1.Cells(3, 2).Value = frmIT_6A_G1.cmb2.Text
            Sheet_IT_6A_G1.Cells(4, 2).Value = frmIT_6A_G1.cmb3.Text
            Sheet_IT_6A_G1.Cells(5, 2).Value = frmIT_6A_G1.cmb4.Text
            Sheet_IT_6A_G1.Cells(6, 2).Value = frmIT_6A_G1.cmb5.Text
            Sheet_IT_6A_G1.Cells(7, 2).Value = frmIT_6A_G1.cmb6.Text
            Sheet_IT_6A_G1.Cells(8, 2).Value = frmIT_6A_G1.cmb7.Text
            Sheet_IT_6A_G1.Cells(9, 2).Value = frmIT_6A_G1.cmb8.Text
            Sheet_IT_6A_G1.Cells(10, 2).Value = frmIT_6A_G1.cmb9.Text
            Sheet_IT_6A_G1.Cells(11, 2).Value = frmIT_6A_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_6A_G1.Cells(2, 3).Value = frmIT_6A_G1.cmb21.Text
            Sheet_IT_6A_G1.Cells(3, 3).Value = frmIT_6A_G1.cmb22.Text
            Sheet_IT_6A_G1.Cells(4, 3).Value = frmIT_6A_G1.cmb23.Text
            Sheet_IT_6A_G1.Cells(5, 3).Value = frmIT_6A_G1.cmb24.Text
            Sheet_IT_6A_G1.Cells(6, 3).Value = frmIT_6A_G1.cmb25.Text
            Sheet_IT_6A_G1.Cells(7, 3).Value = frmIT_6A_G1.cmb26.Text
            Sheet_IT_6A_G1.Cells(8, 3).Value = frmIT_6A_G1.cmb27.Text
            Sheet_IT_6A_G1.Cells(9, 3).Value = frmIT_6A_G1.cmb28.Text
            Sheet_IT_6A_G1.Cells(10, 3).Value = frmIT_6A_G1.cmb29.Text
            Sheet_IT_6A_G1.Cells(11, 3).Value = frmIT_6A_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_6A_G1.Cells(2, 4).Value = frmIT_6A_G1.cmb41.Text
            Sheet_IT_6A_G1.Cells(3, 4).Value = frmIT_6A_G1.cmb42.Text
            Sheet_IT_6A_G1.Cells(4, 4).Value = frmIT_6A_G1.cmb43.Text
            Sheet_IT_6A_G1.Cells(5, 4).Value = frmIT_6A_G1.cmb44.Text
            Sheet_IT_6A_G1.Cells(6, 4).Value = frmIT_6A_G1.cmb45.Text
            Sheet_IT_6A_G1.Cells(7, 4).Value = frmIT_6A_G1.cmb46.Text
            Sheet_IT_6A_G1.Cells(8, 4).Value = frmIT_6A_G1.cmb47.Text
            Sheet_IT_6A_G1.Cells(9, 4).Value = frmIT_6A_G1.cmb48.Text
            Sheet_IT_6A_G1.Cells(10, 4).Value = frmIT_6A_G1.cmb49.Text
            Sheet_IT_6A_G1.Cells(11, 4).Value = frmIT_6A_G1.cmb50.Text

            ' Thursday

            Sheet_IT_6A_G1.Cells(2, 5).Value = frmIT_6A_G1.cmb61.Text
            Sheet_IT_6A_G1.Cells(3, 5).Value = frmIT_6A_G1.cmb62.Text
            Sheet_IT_6A_G1.Cells(4, 5).Value = frmIT_6A_G1.cmb63.Text
            Sheet_IT_6A_G1.Cells(5, 5).Value = frmIT_6A_G1.cmb64.Text
            Sheet_IT_6A_G1.Cells(6, 5).Value = frmIT_6A_G1.cmb65.Text
            Sheet_IT_6A_G1.Cells(7, 5).Value = frmIT_6A_G1.cmb66.Text
            Sheet_IT_6A_G1.Cells(8, 5).Value = frmIT_6A_G1.cmb67.Text
            Sheet_IT_6A_G1.Cells(9, 5).Value = frmIT_6A_G1.cmb68.Text
            Sheet_IT_6A_G1.Cells(10, 5).Value = frmIT_6A_G1.cmb69.Text
            Sheet_IT_6A_G1.Cells(11, 5).Value = frmIT_6A_G1.cmb70.Text

            ' Friday

            Sheet_IT_6A_G1.Cells(2, 6).Value = frmIT_6A_G1.cmb81.Text
            Sheet_IT_6A_G1.Cells(3, 6).Value = frmIT_6A_G1.cmb82.Text
            Sheet_IT_6A_G1.Cells(4, 6).Value = frmIT_6A_G1.cmb83.Text
            Sheet_IT_6A_G1.Cells(5, 6).Value = frmIT_6A_G1.cmb84.Text
            Sheet_IT_6A_G1.Cells(6, 6).Value = frmIT_6A_G1.cmb85.Text
            Sheet_IT_6A_G1.Cells(7, 6).Value = frmIT_6A_G1.cmb86.Text
            Sheet_IT_6A_G1.Cells(8, 6).Value = frmIT_6A_G1.cmb87.Text
            Sheet_IT_6A_G1.Cells(9, 6).Value = frmIT_6A_G1.cmb88.Text
            Sheet_IT_6A_G1.Cells(10, 6).Value = frmIT_6A_G1.cmb89.Text
            Sheet_IT_6A_G1.Cells(11, 6).Value = frmIT_6A_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_6A_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_6A_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_6A_G2.Cells(1, 1).Value = ""
            Sheet_IT_6A_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_6A_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_6A_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_6A_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_6A_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_6A_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_6A_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_6A_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_6A_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_6A_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_6A_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_6A_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_6A_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_6A_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_6A_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_6A_G2.Cells(2, 2).Value = frmIT_6A_G2.cmb1.Text
            Sheet_IT_6A_G2.Cells(3, 2).Value = frmIT_6A_G2.cmb2.Text
            Sheet_IT_6A_G2.Cells(4, 2).Value = frmIT_6A_G2.cmb3.Text
            Sheet_IT_6A_G2.Cells(5, 2).Value = frmIT_6A_G2.cmb4.Text
            Sheet_IT_6A_G2.Cells(6, 2).Value = frmIT_6A_G2.cmb5.Text
            Sheet_IT_6A_G2.Cells(7, 2).Value = frmIT_6A_G2.cmb6.Text
            Sheet_IT_6A_G2.Cells(8, 2).Value = frmIT_6A_G2.cmb7.Text
            Sheet_IT_6A_G2.Cells(9, 2).Value = frmIT_6A_G2.cmb8.Text
            Sheet_IT_6A_G2.Cells(10, 2).Value = frmIT_6A_G2.cmb9.Text
            Sheet_IT_6A_G2.Cells(11, 2).Value = frmIT_6A_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_6A_G2.Cells(2, 3).Value = frmIT_6A_G2.cmb21.Text
            Sheet_IT_6A_G2.Cells(3, 3).Value = frmIT_6A_G2.cmb22.Text
            Sheet_IT_6A_G2.Cells(4, 3).Value = frmIT_6A_G2.cmb23.Text
            Sheet_IT_6A_G2.Cells(5, 3).Value = frmIT_6A_G2.cmb24.Text
            Sheet_IT_6A_G2.Cells(6, 3).Value = frmIT_6A_G2.cmb25.Text
            Sheet_IT_6A_G2.Cells(7, 3).Value = frmIT_6A_G2.cmb26.Text
            Sheet_IT_6A_G2.Cells(8, 3).Value = frmIT_6A_G2.cmb27.Text
            Sheet_IT_6A_G2.Cells(9, 3).Value = frmIT_6A_G2.cmb28.Text
            Sheet_IT_6A_G2.Cells(10, 3).Value = frmIT_6A_G2.cmb29.Text
            Sheet_IT_6A_G2.Cells(11, 3).Value = frmIT_6A_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_6A_G2.Cells(2, 4).Value = frmIT_6A_G2.cmb41.Text
            Sheet_IT_6A_G2.Cells(3, 4).Value = frmIT_6A_G2.cmb42.Text
            Sheet_IT_6A_G2.Cells(4, 4).Value = frmIT_6A_G2.cmb43.Text
            Sheet_IT_6A_G2.Cells(5, 4).Value = frmIT_6A_G2.cmb44.Text
            Sheet_IT_6A_G2.Cells(6, 4).Value = frmIT_6A_G2.cmb45.Text
            Sheet_IT_6A_G2.Cells(7, 4).Value = frmIT_6A_G2.cmb46.Text
            Sheet_IT_6A_G2.Cells(8, 4).Value = frmIT_6A_G2.cmb47.Text
            Sheet_IT_6A_G2.Cells(9, 4).Value = frmIT_6A_G2.cmb48.Text
            Sheet_IT_6A_G2.Cells(10, 4).Value = frmIT_6A_G2.cmb49.Text
            Sheet_IT_6A_G2.Cells(11, 4).Value = frmIT_6A_G2.cmb50.Text

            ' Thursday

            Sheet_IT_6A_G2.Cells(2, 5).Value = frmIT_6A_G2.cmb61.Text
            Sheet_IT_6A_G2.Cells(3, 5).Value = frmIT_6A_G2.cmb62.Text
            Sheet_IT_6A_G2.Cells(4, 5).Value = frmIT_6A_G2.cmb63.Text
            Sheet_IT_6A_G2.Cells(5, 5).Value = frmIT_6A_G2.cmb64.Text
            Sheet_IT_6A_G2.Cells(6, 5).Value = frmIT_6A_G2.cmb65.Text
            Sheet_IT_6A_G2.Cells(7, 5).Value = frmIT_6A_G2.cmb66.Text
            Sheet_IT_6A_G2.Cells(8, 5).Value = frmIT_6A_G2.cmb67.Text
            Sheet_IT_6A_G2.Cells(9, 5).Value = frmIT_6A_G2.cmb68.Text
            Sheet_IT_6A_G2.Cells(10, 5).Value = frmIT_6A_G2.cmb69.Text
            Sheet_IT_6A_G2.Cells(11, 5).Value = frmIT_6A_G2.cmb70.Text

            ' Friday

            Sheet_IT_6A_G2.Cells(2, 6).Value = frmIT_6A_G2.cmb81.Text
            Sheet_IT_6A_G2.Cells(3, 6).Value = frmIT_6A_G2.cmb82.Text
            Sheet_IT_6A_G2.Cells(4, 6).Value = frmIT_6A_G2.cmb83.Text
            Sheet_IT_6A_G2.Cells(5, 6).Value = frmIT_6A_G2.cmb84.Text
            Sheet_IT_6A_G2.Cells(6, 6).Value = frmIT_6A_G2.cmb85.Text
            Sheet_IT_6A_G2.Cells(7, 6).Value = frmIT_6A_G2.cmb86.Text
            Sheet_IT_6A_G2.Cells(8, 6).Value = frmIT_6A_G2.cmb87.Text
            Sheet_IT_6A_G2.Cells(9, 6).Value = frmIT_6A_G2.cmb88.Text
            Sheet_IT_6A_G2.Cells(10, 6).Value = frmIT_6A_G2.cmb89.Text
            Sheet_IT_6A_G2.Cells(11, 6).Value = frmIT_6A_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_6A_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_6A_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_6A_G3.Cells(1, 1).Value = ""
            Sheet_IT_6A_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_6A_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_6A_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_6A_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_6A_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_6A_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_6A_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_6A_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_6A_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_6A_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_6A_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_6A_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_6A_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_6A_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_6A_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_6A_G3.Cells(2, 2).Value = frmIT_6A_G3.cmb1.Text
            Sheet_IT_6A_G3.Cells(3, 2).Value = frmIT_6A_G3.cmb2.Text
            Sheet_IT_6A_G3.Cells(4, 2).Value = frmIT_6A_G3.cmb3.Text
            Sheet_IT_6A_G3.Cells(5, 2).Value = frmIT_6A_G3.cmb4.Text
            Sheet_IT_6A_G3.Cells(6, 2).Value = frmIT_6A_G3.cmb5.Text
            Sheet_IT_6A_G3.Cells(7, 2).Value = frmIT_6A_G3.cmb6.Text
            Sheet_IT_6A_G3.Cells(8, 2).Value = frmIT_6A_G3.cmb7.Text
            Sheet_IT_6A_G3.Cells(9, 2).Value = frmIT_6A_G3.cmb8.Text
            Sheet_IT_6A_G3.Cells(10, 2).Value = frmIT_6A_G3.cmb9.Text
            Sheet_IT_6A_G3.Cells(11, 2).Value = frmIT_6A_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_6A_G3.Cells(2, 3).Value = frmIT_6A_G3.cmb21.Text
            Sheet_IT_6A_G3.Cells(3, 3).Value = frmIT_6A_G3.cmb22.Text
            Sheet_IT_6A_G3.Cells(4, 3).Value = frmIT_6A_G3.cmb23.Text
            Sheet_IT_6A_G3.Cells(5, 3).Value = frmIT_6A_G3.cmb24.Text
            Sheet_IT_6A_G3.Cells(6, 3).Value = frmIT_6A_G3.cmb25.Text
            Sheet_IT_6A_G3.Cells(7, 3).Value = frmIT_6A_G3.cmb26.Text
            Sheet_IT_6A_G3.Cells(8, 3).Value = frmIT_6A_G3.cmb27.Text
            Sheet_IT_6A_G3.Cells(9, 3).Value = frmIT_6A_G3.cmb28.Text
            Sheet_IT_6A_G3.Cells(10, 3).Value = frmIT_6A_G3.cmb29.Text
            Sheet_IT_6A_G3.Cells(11, 3).Value = frmIT_6A_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_6A_G3.Cells(2, 4).Value = frmIT_6A_G3.cmb41.Text
            Sheet_IT_6A_G3.Cells(3, 4).Value = frmIT_6A_G3.cmb42.Text
            Sheet_IT_6A_G3.Cells(4, 4).Value = frmIT_6A_G3.cmb43.Text
            Sheet_IT_6A_G3.Cells(5, 4).Value = frmIT_6A_G3.cmb44.Text
            Sheet_IT_6A_G3.Cells(6, 4).Value = frmIT_6A_G3.cmb45.Text
            Sheet_IT_6A_G3.Cells(7, 4).Value = frmIT_6A_G3.cmb46.Text
            Sheet_IT_6A_G3.Cells(8, 4).Value = frmIT_6A_G3.cmb47.Text
            Sheet_IT_6A_G3.Cells(9, 4).Value = frmIT_6A_G3.cmb48.Text
            Sheet_IT_6A_G3.Cells(10, 4).Value = frmIT_6A_G3.cmb49.Text
            Sheet_IT_6A_G3.Cells(11, 4).Value = frmIT_6A_G3.cmb50.Text

            ' Thursday

            Sheet_IT_6A_G3.Cells(2, 5).Value = frmIT_6A_G3.cmb61.Text
            Sheet_IT_6A_G3.Cells(3, 5).Value = frmIT_6A_G3.cmb62.Text
            Sheet_IT_6A_G3.Cells(4, 5).Value = frmIT_6A_G3.cmb63.Text
            Sheet_IT_6A_G3.Cells(5, 5).Value = frmIT_6A_G3.cmb64.Text
            Sheet_IT_6A_G3.Cells(6, 5).Value = frmIT_6A_G3.cmb65.Text
            Sheet_IT_6A_G3.Cells(7, 5).Value = frmIT_6A_G3.cmb66.Text
            Sheet_IT_6A_G3.Cells(8, 5).Value = frmIT_6A_G3.cmb67.Text
            Sheet_IT_6A_G3.Cells(9, 5).Value = frmIT_6A_G3.cmb68.Text
            Sheet_IT_6A_G3.Cells(10, 5).Value = frmIT_6A_G3.cmb69.Text
            Sheet_IT_6A_G3.Cells(11, 5).Value = frmIT_6A_G3.cmb70.Text

            ' Friday

            Sheet_IT_6A_G3.Cells(2, 6).Value = frmIT_6A_G3.cmb81.Text
            Sheet_IT_6A_G3.Cells(3, 6).Value = frmIT_6A_G3.cmb82.Text
            Sheet_IT_6A_G3.Cells(4, 6).Value = frmIT_6A_G3.cmb83.Text
            Sheet_IT_6A_G3.Cells(5, 6).Value = frmIT_6A_G3.cmb84.Text
            Sheet_IT_6A_G3.Cells(6, 6).Value = frmIT_6A_G3.cmb85.Text
            Sheet_IT_6A_G3.Cells(7, 6).Value = frmIT_6A_G3.cmb86.Text
            Sheet_IT_6A_G3.Cells(8, 6).Value = frmIT_6A_G3.cmb87.Text
            Sheet_IT_6A_G3.Cells(9, 6).Value = frmIT_6A_G3.cmb88.Text
            Sheet_IT_6A_G3.Cells(10, 6).Value = frmIT_6A_G3.cmb89.Text
            Sheet_IT_6A_G3.Cells(11, 6).Value = frmIT_6A_G3.cmb90.Text
            ''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''
            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_6B_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_6B_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_6B_G1.Cells(1, 1).Value = ""
            Sheet_IT_6B_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_6B_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_6B_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_6B_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_6B_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_6B_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_6B_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_6B_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_6B_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_6B_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_6B_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_6B_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_6B_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_6B_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_6B_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_6B_G1.Cells(2, 2).Value = frmIT_6B_G1.cmb1.Text
            Sheet_IT_6B_G1.Cells(3, 2).Value = frmIT_6B_G1.cmb2.Text
            Sheet_IT_6B_G1.Cells(4, 2).Value = frmIT_6B_G1.cmb3.Text
            Sheet_IT_6B_G1.Cells(5, 2).Value = frmIT_6B_G1.cmb4.Text
            Sheet_IT_6B_G1.Cells(6, 2).Value = frmIT_6B_G1.cmb5.Text
            Sheet_IT_6B_G1.Cells(7, 2).Value = frmIT_6B_G1.cmb6.Text
            Sheet_IT_6B_G1.Cells(8, 2).Value = frmIT_6B_G1.cmb7.Text
            Sheet_IT_6B_G1.Cells(9, 2).Value = frmIT_6B_G1.cmb8.Text
            Sheet_IT_6B_G1.Cells(10, 2).Value = frmIT_6B_G1.cmb9.Text
            Sheet_IT_6B_G1.Cells(11, 2).Value = frmIT_6B_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_6B_G1.Cells(2, 3).Value = frmIT_6B_G1.cmb21.Text
            Sheet_IT_6B_G1.Cells(3, 3).Value = frmIT_6B_G1.cmb22.Text
            Sheet_IT_6B_G1.Cells(4, 3).Value = frmIT_6B_G1.cmb23.Text
            Sheet_IT_6B_G1.Cells(5, 3).Value = frmIT_6B_G1.cmb24.Text
            Sheet_IT_6B_G1.Cells(6, 3).Value = frmIT_6B_G1.cmb25.Text
            Sheet_IT_6B_G1.Cells(7, 3).Value = frmIT_6B_G1.cmb26.Text
            Sheet_IT_6B_G1.Cells(8, 3).Value = frmIT_6B_G1.cmb27.Text
            Sheet_IT_6B_G1.Cells(9, 3).Value = frmIT_6B_G1.cmb28.Text
            Sheet_IT_6B_G1.Cells(10, 3).Value = frmIT_6B_G1.cmb29.Text
            Sheet_IT_6B_G1.Cells(11, 3).Value = frmIT_6B_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_6B_G1.Cells(2, 4).Value = frmIT_6B_G1.cmb41.Text
            Sheet_IT_6B_G1.Cells(3, 4).Value = frmIT_6B_G1.cmb42.Text
            Sheet_IT_6B_G1.Cells(4, 4).Value = frmIT_6B_G1.cmb43.Text
            Sheet_IT_6B_G1.Cells(5, 4).Value = frmIT_6B_G1.cmb44.Text
            Sheet_IT_6B_G1.Cells(6, 4).Value = frmIT_6B_G1.cmb45.Text
            Sheet_IT_6B_G1.Cells(7, 4).Value = frmIT_6B_G1.cmb46.Text
            Sheet_IT_6B_G1.Cells(8, 4).Value = frmIT_6B_G1.cmb47.Text
            Sheet_IT_6B_G1.Cells(9, 4).Value = frmIT_6B_G1.cmb48.Text
            Sheet_IT_6B_G1.Cells(10, 4).Value = frmIT_6B_G1.cmb49.Text
            Sheet_IT_6B_G1.Cells(11, 4).Value = frmIT_6B_G1.cmb50.Text

            ' Thursday

            Sheet_IT_6B_G1.Cells(2, 5).Value = frmIT_6B_G1.cmb61.Text
            Sheet_IT_6B_G1.Cells(3, 5).Value = frmIT_6B_G1.cmb62.Text
            Sheet_IT_6B_G1.Cells(4, 5).Value = frmIT_6B_G1.cmb63.Text
            Sheet_IT_6B_G1.Cells(5, 5).Value = frmIT_6B_G1.cmb64.Text
            Sheet_IT_6B_G1.Cells(6, 5).Value = frmIT_6B_G1.cmb65.Text
            Sheet_IT_6B_G1.Cells(7, 5).Value = frmIT_6B_G1.cmb66.Text
            Sheet_IT_6B_G1.Cells(8, 5).Value = frmIT_6B_G1.cmb67.Text
            Sheet_IT_6B_G1.Cells(9, 5).Value = frmIT_6B_G1.cmb68.Text
            Sheet_IT_6B_G1.Cells(10, 5).Value = frmIT_6B_G1.cmb69.Text
            Sheet_IT_6B_G1.Cells(11, 5).Value = frmIT_6B_G1.cmb70.Text

            ' Friday

            Sheet_IT_6B_G1.Cells(2, 6).Value = frmIT_6B_G1.cmb81.Text
            Sheet_IT_6B_G1.Cells(3, 6).Value = frmIT_6B_G1.cmb82.Text
            Sheet_IT_6B_G1.Cells(4, 6).Value = frmIT_6B_G1.cmb83.Text
            Sheet_IT_6B_G1.Cells(5, 6).Value = frmIT_6B_G1.cmb84.Text
            Sheet_IT_6B_G1.Cells(6, 6).Value = frmIT_6B_G1.cmb85.Text
            Sheet_IT_6B_G1.Cells(7, 6).Value = frmIT_6B_G1.cmb86.Text
            Sheet_IT_6B_G1.Cells(8, 6).Value = frmIT_6B_G1.cmb87.Text
            Sheet_IT_6B_G1.Cells(9, 6).Value = frmIT_6B_G1.cmb88.Text
            Sheet_IT_6B_G1.Cells(10, 6).Value = frmIT_6B_G1.cmb89.Text
            Sheet_IT_6B_G1.Cells(11, 6).Value = frmIT_6B_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_6B_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_6B_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_6B_G2.Cells(1, 1).Value = ""
            Sheet_IT_6B_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_6B_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_6B_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_6B_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_6B_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_6B_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_6B_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_6B_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_6B_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_6B_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_6B_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_6B_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_6B_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_6B_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_6B_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_6B_G2.Cells(2, 2).Value = frmIT_6B_G2.cmb1.Text
            Sheet_IT_6B_G2.Cells(3, 2).Value = frmIT_6B_G2.cmb2.Text
            Sheet_IT_6B_G2.Cells(4, 2).Value = frmIT_6B_G2.cmb3.Text
            Sheet_IT_6B_G2.Cells(5, 2).Value = frmIT_6B_G2.cmb4.Text
            Sheet_IT_6B_G2.Cells(6, 2).Value = frmIT_6B_G2.cmb5.Text
            Sheet_IT_6B_G2.Cells(7, 2).Value = frmIT_6B_G2.cmb6.Text
            Sheet_IT_6B_G2.Cells(8, 2).Value = frmIT_6B_G2.cmb7.Text
            Sheet_IT_6B_G2.Cells(9, 2).Value = frmIT_6B_G2.cmb8.Text
            Sheet_IT_6B_G2.Cells(10, 2).Value = frmIT_6B_G2.cmb9.Text
            Sheet_IT_6B_G2.Cells(11, 2).Value = frmIT_6B_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_6B_G2.Cells(2, 3).Value = frmIT_6B_G2.cmb21.Text
            Sheet_IT_6B_G2.Cells(3, 3).Value = frmIT_6B_G2.cmb22.Text
            Sheet_IT_6B_G2.Cells(4, 3).Value = frmIT_6B_G2.cmb23.Text
            Sheet_IT_6B_G2.Cells(5, 3).Value = frmIT_6B_G2.cmb24.Text
            Sheet_IT_6B_G2.Cells(6, 3).Value = frmIT_6B_G2.cmb25.Text
            Sheet_IT_6B_G2.Cells(7, 3).Value = frmIT_6B_G2.cmb26.Text
            Sheet_IT_6B_G2.Cells(8, 3).Value = frmIT_6B_G2.cmb27.Text
            Sheet_IT_6B_G2.Cells(9, 3).Value = frmIT_6B_G2.cmb28.Text
            Sheet_IT_6B_G2.Cells(10, 3).Value = frmIT_6B_G2.cmb29.Text
            Sheet_IT_6B_G2.Cells(11, 3).Value = frmIT_6B_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_6B_G2.Cells(2, 4).Value = frmIT_6B_G2.cmb41.Text
            Sheet_IT_6B_G2.Cells(3, 4).Value = frmIT_6B_G2.cmb42.Text
            Sheet_IT_6B_G2.Cells(4, 4).Value = frmIT_6B_G2.cmb43.Text
            Sheet_IT_6B_G2.Cells(5, 4).Value = frmIT_6B_G2.cmb44.Text
            Sheet_IT_6B_G2.Cells(6, 4).Value = frmIT_6B_G2.cmb45.Text
            Sheet_IT_6B_G2.Cells(7, 4).Value = frmIT_6B_G2.cmb46.Text
            Sheet_IT_6B_G2.Cells(8, 4).Value = frmIT_6B_G2.cmb47.Text
            Sheet_IT_6B_G2.Cells(9, 4).Value = frmIT_6B_G2.cmb48.Text
            Sheet_IT_6B_G2.Cells(10, 4).Value = frmIT_6B_G2.cmb49.Text
            Sheet_IT_6B_G2.Cells(11, 4).Value = frmIT_6B_G2.cmb50.Text

            ' Thursday

            Sheet_IT_6B_G2.Cells(2, 5).Value = frmIT_6B_G2.cmb61.Text
            Sheet_IT_6B_G2.Cells(3, 5).Value = frmIT_6B_G2.cmb62.Text
            Sheet_IT_6B_G2.Cells(4, 5).Value = frmIT_6B_G2.cmb63.Text
            Sheet_IT_6B_G2.Cells(5, 5).Value = frmIT_6B_G2.cmb64.Text
            Sheet_IT_6B_G2.Cells(6, 5).Value = frmIT_6B_G2.cmb65.Text
            Sheet_IT_6B_G2.Cells(7, 5).Value = frmIT_6B_G2.cmb66.Text
            Sheet_IT_6B_G2.Cells(8, 5).Value = frmIT_6B_G2.cmb67.Text
            Sheet_IT_6B_G2.Cells(9, 5).Value = frmIT_6B_G2.cmb68.Text
            Sheet_IT_6B_G2.Cells(10, 5).Value = frmIT_6B_G2.cmb69.Text
            Sheet_IT_6B_G2.Cells(11, 5).Value = frmIT_6B_G2.cmb70.Text

            ' Friday

            Sheet_IT_6B_G2.Cells(2, 6).Value = frmIT_6B_G2.cmb81.Text
            Sheet_IT_6B_G2.Cells(3, 6).Value = frmIT_6B_G2.cmb82.Text
            Sheet_IT_6B_G2.Cells(4, 6).Value = frmIT_6B_G2.cmb83.Text
            Sheet_IT_6B_G2.Cells(5, 6).Value = frmIT_6B_G2.cmb84.Text
            Sheet_IT_6B_G2.Cells(6, 6).Value = frmIT_6B_G2.cmb85.Text
            Sheet_IT_6B_G2.Cells(7, 6).Value = frmIT_6B_G2.cmb86.Text
            Sheet_IT_6B_G2.Cells(8, 6).Value = frmIT_6B_G2.cmb87.Text
            Sheet_IT_6B_G2.Cells(9, 6).Value = frmIT_6B_G2.cmb88.Text
            Sheet_IT_6B_G2.Cells(10, 6).Value = frmIT_6B_G2.cmb89.Text
            Sheet_IT_6B_G2.Cells(11, 6).Value = frmIT_6B_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_6B_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_6B_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_6B_G3.Cells(1, 1).Value = ""
            Sheet_IT_6B_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_6B_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_6B_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_6B_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_6B_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_6B_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_6B_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_6B_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_6B_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_6B_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_6B_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_6B_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_6B_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_6B_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_6B_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_6B_G3.Cells(2, 2).Value = frmIT_6B_G3.cmb1.Text
            Sheet_IT_6B_G3.Cells(3, 2).Value = frmIT_6B_G3.cmb2.Text
            Sheet_IT_6B_G3.Cells(4, 2).Value = frmIT_6B_G3.cmb3.Text
            Sheet_IT_6B_G3.Cells(5, 2).Value = frmIT_6B_G3.cmb4.Text
            Sheet_IT_6B_G3.Cells(6, 2).Value = frmIT_6B_G3.cmb5.Text
            Sheet_IT_6B_G3.Cells(7, 2).Value = frmIT_6B_G3.cmb6.Text
            Sheet_IT_6B_G3.Cells(8, 2).Value = frmIT_6B_G3.cmb7.Text
            Sheet_IT_6B_G3.Cells(9, 2).Value = frmIT_6B_G3.cmb8.Text
            Sheet_IT_6B_G3.Cells(10, 2).Value = frmIT_6B_G3.cmb9.Text
            Sheet_IT_6B_G3.Cells(11, 2).Value = frmIT_6B_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_6B_G3.Cells(2, 3).Value = frmIT_6B_G3.cmb21.Text
            Sheet_IT_6B_G3.Cells(3, 3).Value = frmIT_6B_G3.cmb22.Text
            Sheet_IT_6B_G3.Cells(4, 3).Value = frmIT_6B_G3.cmb23.Text
            Sheet_IT_6B_G3.Cells(5, 3).Value = frmIT_6B_G3.cmb24.Text
            Sheet_IT_6B_G3.Cells(6, 3).Value = frmIT_6B_G3.cmb25.Text
            Sheet_IT_6B_G3.Cells(7, 3).Value = frmIT_6B_G3.cmb26.Text
            Sheet_IT_6B_G3.Cells(8, 3).Value = frmIT_6B_G3.cmb27.Text
            Sheet_IT_6B_G3.Cells(9, 3).Value = frmIT_6B_G3.cmb28.Text
            Sheet_IT_6B_G3.Cells(10, 3).Value = frmIT_6B_G3.cmb29.Text
            Sheet_IT_6B_G3.Cells(11, 3).Value = frmIT_6B_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_6B_G3.Cells(2, 4).Value = frmIT_6B_G3.cmb41.Text
            Sheet_IT_6B_G3.Cells(3, 4).Value = frmIT_6B_G3.cmb42.Text
            Sheet_IT_6B_G3.Cells(4, 4).Value = frmIT_6B_G3.cmb43.Text
            Sheet_IT_6B_G3.Cells(5, 4).Value = frmIT_6B_G3.cmb44.Text
            Sheet_IT_6B_G3.Cells(6, 4).Value = frmIT_6B_G3.cmb45.Text
            Sheet_IT_6B_G3.Cells(7, 4).Value = frmIT_6B_G3.cmb46.Text
            Sheet_IT_6B_G3.Cells(8, 4).Value = frmIT_6B_G3.cmb47.Text
            Sheet_IT_6B_G3.Cells(9, 4).Value = frmIT_6B_G3.cmb48.Text
            Sheet_IT_6B_G3.Cells(10, 4).Value = frmIT_6B_G3.cmb49.Text
            Sheet_IT_6B_G3.Cells(11, 4).Value = frmIT_6B_G3.cmb50.Text

            ' Thursday

            Sheet_IT_6B_G3.Cells(2, 5).Value = frmIT_6B_G3.cmb61.Text
            Sheet_IT_6B_G3.Cells(3, 5).Value = frmIT_6B_G3.cmb62.Text
            Sheet_IT_6B_G3.Cells(4, 5).Value = frmIT_6B_G3.cmb63.Text
            Sheet_IT_6B_G3.Cells(5, 5).Value = frmIT_6B_G3.cmb64.Text
            Sheet_IT_6B_G3.Cells(6, 5).Value = frmIT_6B_G3.cmb65.Text
            Sheet_IT_6B_G3.Cells(7, 5).Value = frmIT_6B_G3.cmb66.Text
            Sheet_IT_6B_G3.Cells(8, 5).Value = frmIT_6B_G3.cmb67.Text
            Sheet_IT_6B_G3.Cells(9, 5).Value = frmIT_6B_G3.cmb68.Text
            Sheet_IT_6B_G3.Cells(10, 5).Value = frmIT_6B_G3.cmb69.Text
            Sheet_IT_6B_G3.Cells(11, 5).Value = frmIT_6B_G3.cmb70.Text

            ' Friday

            Sheet_IT_6B_G3.Cells(2, 6).Value = frmIT_6B_G3.cmb81.Text
            Sheet_IT_6B_G3.Cells(3, 6).Value = frmIT_6B_G3.cmb82.Text
            Sheet_IT_6B_G3.Cells(4, 6).Value = frmIT_6B_G3.cmb83.Text
            Sheet_IT_6B_G3.Cells(5, 6).Value = frmIT_6B_G3.cmb84.Text
            Sheet_IT_6B_G3.Cells(6, 6).Value = frmIT_6B_G3.cmb85.Text
            Sheet_IT_6B_G3.Cells(7, 6).Value = frmIT_6B_G3.cmb86.Text
            Sheet_IT_6B_G3.Cells(8, 6).Value = frmIT_6B_G3.cmb87.Text
            Sheet_IT_6B_G3.Cells(9, 6).Value = frmIT_6B_G3.cmb88.Text
            Sheet_IT_6B_G3.Cells(10, 6).Value = frmIT_6B_G3.cmb89.Text
            Sheet_IT_6B_G3.Cells(11, 6).Value = frmIT_6B_G3.cmb90.Text

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4A_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4A_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4A_G1.Cells(1, 1).Value = ""
            Sheet_IT_4A_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4A_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4A_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4A_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4A_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4A_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4A_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4A_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4A_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4A_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4A_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_4A_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4A_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4A_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4A_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4A_G1.Cells(2, 2).Value = frmIT_4A_G1.cmb1.Text
            Sheet_IT_4A_G1.Cells(3, 2).Value = frmIT_4A_G1.cmb2.Text
            Sheet_IT_4A_G1.Cells(4, 2).Value = frmIT_4A_G1.cmb3.Text
            Sheet_IT_4A_G1.Cells(5, 2).Value = frmIT_4A_G1.cmb4.Text
            Sheet_IT_4A_G1.Cells(6, 2).Value = frmIT_4A_G1.cmb5.Text
            Sheet_IT_4A_G1.Cells(7, 2).Value = frmIT_4A_G1.cmb6.Text
            Sheet_IT_4A_G1.Cells(8, 2).Value = frmIT_4A_G1.cmb7.Text
            Sheet_IT_4A_G1.Cells(9, 2).Value = frmIT_4A_G1.cmb8.Text
            Sheet_IT_4A_G1.Cells(10, 2).Value = frmIT_4A_G1.cmb9.Text
            Sheet_IT_4A_G1.Cells(11, 2).Value = frmIT_4A_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_4A_G1.Cells(2, 3).Value = frmIT_4A_G1.cmb21.Text
            Sheet_IT_4A_G1.Cells(3, 3).Value = frmIT_4A_G1.cmb22.Text
            Sheet_IT_4A_G1.Cells(4, 3).Value = frmIT_4A_G1.cmb23.Text
            Sheet_IT_4A_G1.Cells(5, 3).Value = frmIT_4A_G1.cmb24.Text
            Sheet_IT_4A_G1.Cells(6, 3).Value = frmIT_4A_G1.cmb25.Text
            Sheet_IT_4A_G1.Cells(7, 3).Value = frmIT_4A_G1.cmb26.Text
            Sheet_IT_4A_G1.Cells(8, 3).Value = frmIT_4A_G1.cmb27.Text
            Sheet_IT_4A_G1.Cells(9, 3).Value = frmIT_4A_G1.cmb28.Text
            Sheet_IT_4A_G1.Cells(10, 3).Value = frmIT_4A_G1.cmb29.Text
            Sheet_IT_4A_G1.Cells(11, 3).Value = frmIT_4A_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_4A_G1.Cells(2, 4).Value = frmIT_4A_G1.cmb41.Text
            Sheet_IT_4A_G1.Cells(3, 4).Value = frmIT_4A_G1.cmb42.Text
            Sheet_IT_4A_G1.Cells(4, 4).Value = frmIT_4A_G1.cmb43.Text
            Sheet_IT_4A_G1.Cells(5, 4).Value = frmIT_4A_G1.cmb44.Text
            Sheet_IT_4A_G1.Cells(6, 4).Value = frmIT_4A_G1.cmb45.Text
            Sheet_IT_4A_G1.Cells(7, 4).Value = frmIT_4A_G1.cmb46.Text
            Sheet_IT_4A_G1.Cells(8, 4).Value = frmIT_4A_G1.cmb47.Text
            Sheet_IT_4A_G1.Cells(9, 4).Value = frmIT_4A_G1.cmb48.Text
            Sheet_IT_4A_G1.Cells(10, 4).Value = frmIT_4A_G1.cmb49.Text
            Sheet_IT_4A_G1.Cells(11, 4).Value = frmIT_4A_G1.cmb50.Text

            ' Thursday

            Sheet_IT_4A_G1.Cells(2, 5).Value = frmIT_4A_G1.cmb61.Text
            Sheet_IT_4A_G1.Cells(3, 5).Value = frmIT_4A_G1.cmb62.Text
            Sheet_IT_4A_G1.Cells(4, 5).Value = frmIT_4A_G1.cmb63.Text
            Sheet_IT_4A_G1.Cells(5, 5).Value = frmIT_4A_G1.cmb64.Text
            Sheet_IT_4A_G1.Cells(6, 5).Value = frmIT_4A_G1.cmb65.Text
            Sheet_IT_4A_G1.Cells(7, 5).Value = frmIT_4A_G1.cmb66.Text
            Sheet_IT_4A_G1.Cells(8, 5).Value = frmIT_4A_G1.cmb67.Text
            Sheet_IT_4A_G1.Cells(9, 5).Value = frmIT_4A_G1.cmb68.Text
            Sheet_IT_4A_G1.Cells(10, 5).Value = frmIT_4A_G1.cmb69.Text
            Sheet_IT_4A_G1.Cells(11, 5).Value = frmIT_4A_G1.cmb70.Text

            ' Friday

            Sheet_IT_4A_G1.Cells(2, 6).Value = frmIT_4A_G1.cmb81.Text
            Sheet_IT_4A_G1.Cells(3, 6).Value = frmIT_4A_G1.cmb82.Text
            Sheet_IT_4A_G1.Cells(4, 6).Value = frmIT_4A_G1.cmb83.Text
            Sheet_IT_4A_G1.Cells(5, 6).Value = frmIT_4A_G1.cmb84.Text
            Sheet_IT_4A_G1.Cells(6, 6).Value = frmIT_4A_G1.cmb85.Text
            Sheet_IT_4A_G1.Cells(7, 6).Value = frmIT_4A_G1.cmb86.Text
            Sheet_IT_4A_G1.Cells(8, 6).Value = frmIT_4A_G1.cmb87.Text
            Sheet_IT_4A_G1.Cells(9, 6).Value = frmIT_4A_G1.cmb88.Text
            Sheet_IT_4A_G1.Cells(10, 6).Value = frmIT_4A_G1.cmb89.Text
            Sheet_IT_4A_G1.Cells(11, 6).Value = frmIT_4A_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4A_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4A_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4A_G2.Cells(1, 1).Value = ""
            Sheet_IT_4A_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4A_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4A_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4A_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4A_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4A_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4A_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4A_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4A_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4A_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4A_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_4A_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4A_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4A_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4A_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4A_G2.Cells(2, 2).Value = frmIT_4A_G2.cmb1.Text
            Sheet_IT_4A_G2.Cells(3, 2).Value = frmIT_4A_G2.cmb2.Text
            Sheet_IT_4A_G2.Cells(4, 2).Value = frmIT_4A_G2.cmb3.Text
            Sheet_IT_4A_G2.Cells(5, 2).Value = frmIT_4A_G2.cmb4.Text
            Sheet_IT_4A_G2.Cells(6, 2).Value = frmIT_4A_G2.cmb5.Text
            Sheet_IT_4A_G2.Cells(7, 2).Value = frmIT_4A_G2.cmb6.Text
            Sheet_IT_4A_G2.Cells(8, 2).Value = frmIT_4A_G2.cmb7.Text
            Sheet_IT_4A_G2.Cells(9, 2).Value = frmIT_4A_G2.cmb8.Text
            Sheet_IT_4A_G2.Cells(10, 2).Value = frmIT_4A_G2.cmb9.Text
            Sheet_IT_4A_G2.Cells(11, 2).Value = frmIT_4A_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_4A_G2.Cells(2, 3).Value = frmIT_4A_G2.cmb21.Text
            Sheet_IT_4A_G2.Cells(3, 3).Value = frmIT_4A_G2.cmb22.Text
            Sheet_IT_4A_G2.Cells(4, 3).Value = frmIT_4A_G2.cmb23.Text
            Sheet_IT_4A_G2.Cells(5, 3).Value = frmIT_4A_G2.cmb24.Text
            Sheet_IT_4A_G2.Cells(6, 3).Value = frmIT_4A_G2.cmb25.Text
            Sheet_IT_4A_G2.Cells(7, 3).Value = frmIT_4A_G2.cmb26.Text
            Sheet_IT_4A_G2.Cells(8, 3).Value = frmIT_4A_G2.cmb27.Text
            Sheet_IT_4A_G2.Cells(9, 3).Value = frmIT_4A_G2.cmb28.Text
            Sheet_IT_4A_G2.Cells(10, 3).Value = frmIT_4A_G2.cmb29.Text
            Sheet_IT_4A_G2.Cells(11, 3).Value = frmIT_4A_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_4A_G2.Cells(2, 4).Value = frmIT_4A_G2.cmb41.Text
            Sheet_IT_4A_G2.Cells(3, 4).Value = frmIT_4A_G2.cmb42.Text
            Sheet_IT_4A_G2.Cells(4, 4).Value = frmIT_4A_G2.cmb43.Text
            Sheet_IT_4A_G2.Cells(5, 4).Value = frmIT_4A_G2.cmb44.Text
            Sheet_IT_4A_G2.Cells(6, 4).Value = frmIT_4A_G2.cmb45.Text
            Sheet_IT_4A_G2.Cells(7, 4).Value = frmIT_4A_G2.cmb46.Text
            Sheet_IT_4A_G2.Cells(8, 4).Value = frmIT_4A_G2.cmb47.Text
            Sheet_IT_4A_G2.Cells(9, 4).Value = frmIT_4A_G2.cmb48.Text
            Sheet_IT_4A_G2.Cells(10, 4).Value = frmIT_4A_G2.cmb49.Text
            Sheet_IT_4A_G2.Cells(11, 4).Value = frmIT_4A_G2.cmb50.Text

            ' Thursday

            Sheet_IT_4A_G2.Cells(2, 5).Value = frmIT_4A_G2.cmb61.Text
            Sheet_IT_4A_G2.Cells(3, 5).Value = frmIT_4A_G2.cmb62.Text
            Sheet_IT_4A_G2.Cells(4, 5).Value = frmIT_4A_G2.cmb63.Text
            Sheet_IT_4A_G2.Cells(5, 5).Value = frmIT_4A_G2.cmb64.Text
            Sheet_IT_4A_G2.Cells(6, 5).Value = frmIT_4A_G2.cmb65.Text
            Sheet_IT_4A_G2.Cells(7, 5).Value = frmIT_4A_G2.cmb66.Text
            Sheet_IT_4A_G2.Cells(8, 5).Value = frmIT_4A_G2.cmb67.Text
            Sheet_IT_4A_G2.Cells(9, 5).Value = frmIT_4A_G2.cmb68.Text
            Sheet_IT_4A_G2.Cells(10, 5).Value = frmIT_4A_G2.cmb69.Text
            Sheet_IT_4A_G2.Cells(11, 5).Value = frmIT_4A_G2.cmb70.Text

            ' Friday

            Sheet_IT_4A_G2.Cells(2, 6).Value = frmIT_4A_G2.cmb81.Text
            Sheet_IT_4A_G2.Cells(3, 6).Value = frmIT_4A_G2.cmb82.Text
            Sheet_IT_4A_G2.Cells(4, 6).Value = frmIT_4A_G2.cmb83.Text
            Sheet_IT_4A_G2.Cells(5, 6).Value = frmIT_4A_G2.cmb84.Text
            Sheet_IT_4A_G2.Cells(6, 6).Value = frmIT_4A_G2.cmb85.Text
            Sheet_IT_4A_G2.Cells(7, 6).Value = frmIT_4A_G2.cmb86.Text
            Sheet_IT_4A_G2.Cells(8, 6).Value = frmIT_4A_G2.cmb87.Text
            Sheet_IT_4A_G2.Cells(9, 6).Value = frmIT_4A_G2.cmb88.Text
            Sheet_IT_4A_G2.Cells(10, 6).Value = frmIT_4A_G2.cmb89.Text
            Sheet_IT_4A_G2.Cells(11, 6).Value = frmIT_4A_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4A_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4A_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4A_G3.Cells(1, 1).Value = ""
            Sheet_IT_4A_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4A_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4A_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4A_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4A_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4A_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4A_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4A_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4A_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4A_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4A_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_4A_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4A_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4A_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4A_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4A_G3.Cells(2, 2).Value = frmIT_4A_G3.cmb1.Text
            Sheet_IT_4A_G3.Cells(3, 2).Value = frmIT_4A_G3.cmb2.Text
            Sheet_IT_4A_G3.Cells(4, 2).Value = frmIT_4A_G3.cmb3.Text
            Sheet_IT_4A_G3.Cells(5, 2).Value = frmIT_4A_G3.cmb4.Text
            Sheet_IT_4A_G3.Cells(6, 2).Value = frmIT_4A_G3.cmb5.Text
            Sheet_IT_4A_G3.Cells(7, 2).Value = frmIT_4A_G3.cmb6.Text
            Sheet_IT_4A_G3.Cells(8, 2).Value = frmIT_4A_G3.cmb7.Text
            Sheet_IT_4A_G3.Cells(9, 2).Value = frmIT_4A_G3.cmb8.Text
            Sheet_IT_4A_G3.Cells(10, 2).Value = frmIT_4A_G3.cmb9.Text
            Sheet_IT_4A_G3.Cells(11, 2).Value = frmIT_4A_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_4A_G3.Cells(2, 3).Value = frmIT_4A_G3.cmb21.Text
            Sheet_IT_4A_G3.Cells(3, 3).Value = frmIT_4A_G3.cmb22.Text
            Sheet_IT_4A_G3.Cells(4, 3).Value = frmIT_4A_G3.cmb23.Text
            Sheet_IT_4A_G3.Cells(5, 3).Value = frmIT_4A_G3.cmb24.Text
            Sheet_IT_4A_G3.Cells(6, 3).Value = frmIT_4A_G3.cmb25.Text
            Sheet_IT_4A_G3.Cells(7, 3).Value = frmIT_4A_G3.cmb26.Text
            Sheet_IT_4A_G3.Cells(8, 3).Value = frmIT_4A_G3.cmb27.Text
            Sheet_IT_4A_G3.Cells(9, 3).Value = frmIT_4A_G3.cmb28.Text
            Sheet_IT_4A_G3.Cells(10, 3).Value = frmIT_4A_G3.cmb29.Text
            Sheet_IT_4A_G3.Cells(11, 3).Value = frmIT_4A_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_4A_G3.Cells(2, 4).Value = frmIT_4A_G3.cmb41.Text
            Sheet_IT_4A_G3.Cells(3, 4).Value = frmIT_4A_G3.cmb42.Text
            Sheet_IT_4A_G3.Cells(4, 4).Value = frmIT_4A_G3.cmb43.Text
            Sheet_IT_4A_G3.Cells(5, 4).Value = frmIT_4A_G3.cmb44.Text
            Sheet_IT_4A_G3.Cells(6, 4).Value = frmIT_4A_G3.cmb45.Text
            Sheet_IT_4A_G3.Cells(7, 4).Value = frmIT_4A_G3.cmb46.Text
            Sheet_IT_4A_G3.Cells(8, 4).Value = frmIT_4A_G3.cmb47.Text
            Sheet_IT_4A_G3.Cells(9, 4).Value = frmIT_4A_G3.cmb48.Text
            Sheet_IT_4A_G3.Cells(10, 4).Value = frmIT_4A_G3.cmb49.Text
            Sheet_IT_4A_G3.Cells(11, 4).Value = frmIT_4A_G3.cmb50.Text

            ' Thursday

            Sheet_IT_4A_G3.Cells(2, 5).Value = frmIT_4A_G3.cmb61.Text
            Sheet_IT_4A_G3.Cells(3, 5).Value = frmIT_4A_G3.cmb62.Text
            Sheet_IT_4A_G3.Cells(4, 5).Value = frmIT_4A_G3.cmb63.Text
            Sheet_IT_4A_G3.Cells(5, 5).Value = frmIT_4A_G3.cmb64.Text
            Sheet_IT_4A_G3.Cells(6, 5).Value = frmIT_4A_G3.cmb65.Text
            Sheet_IT_4A_G3.Cells(7, 5).Value = frmIT_4A_G3.cmb66.Text
            Sheet_IT_4A_G3.Cells(8, 5).Value = frmIT_4A_G3.cmb67.Text
            Sheet_IT_4A_G3.Cells(9, 5).Value = frmIT_4A_G3.cmb68.Text
            Sheet_IT_4A_G3.Cells(10, 5).Value = frmIT_4A_G3.cmb69.Text
            Sheet_IT_4A_G3.Cells(11, 5).Value = frmIT_4A_G3.cmb70.Text

            ' Friday

            Sheet_IT_4A_G3.Cells(2, 6).Value = frmIT_4A_G3.cmb81.Text
            Sheet_IT_4A_G3.Cells(3, 6).Value = frmIT_4A_G3.cmb82.Text
            Sheet_IT_4A_G3.Cells(4, 6).Value = frmIT_4A_G3.cmb83.Text
            Sheet_IT_4A_G3.Cells(5, 6).Value = frmIT_4A_G3.cmb84.Text
            Sheet_IT_4A_G3.Cells(6, 6).Value = frmIT_4A_G3.cmb85.Text
            Sheet_IT_4A_G3.Cells(7, 6).Value = frmIT_4A_G3.cmb86.Text
            Sheet_IT_4A_G3.Cells(8, 6).Value = frmIT_4A_G3.cmb87.Text
            Sheet_IT_4A_G3.Cells(9, 6).Value = frmIT_4A_G3.cmb88.Text
            Sheet_IT_4A_G3.Cells(10, 6).Value = frmIT_4A_G3.cmb89.Text
            Sheet_IT_4A_G3.Cells(11, 6).Value = frmIT_4A_G3.cmb90.Text
            ''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''
            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4B_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4B_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4B_G1.Cells(1, 1).Value = ""
            Sheet_IT_4B_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4B_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4B_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4B_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4B_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4B_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4B_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4B_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4B_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4B_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4B_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_4B_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4B_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4B_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4B_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4B_G1.Cells(2, 2).Value = frmIT_4B_G1.cmb1.Text
            Sheet_IT_4B_G1.Cells(3, 2).Value = frmIT_4B_G1.cmb2.Text
            Sheet_IT_4B_G1.Cells(4, 2).Value = frmIT_4B_G1.cmb3.Text
            Sheet_IT_4B_G1.Cells(5, 2).Value = frmIT_4B_G1.cmb4.Text
            Sheet_IT_4B_G1.Cells(6, 2).Value = frmIT_4B_G1.cmb5.Text
            Sheet_IT_4B_G1.Cells(7, 2).Value = frmIT_4B_G1.cmb6.Text
            Sheet_IT_4B_G1.Cells(8, 2).Value = frmIT_4B_G1.cmb7.Text
            Sheet_IT_4B_G1.Cells(9, 2).Value = frmIT_4B_G1.cmb8.Text
            Sheet_IT_4B_G1.Cells(10, 2).Value = frmIT_4B_G1.cmb9.Text
            Sheet_IT_4B_G1.Cells(11, 2).Value = frmIT_4B_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_4B_G1.Cells(2, 3).Value = frmIT_4B_G1.cmb21.Text
            Sheet_IT_4B_G1.Cells(3, 3).Value = frmIT_4B_G1.cmb22.Text
            Sheet_IT_4B_G1.Cells(4, 3).Value = frmIT_4B_G1.cmb23.Text
            Sheet_IT_4B_G1.Cells(5, 3).Value = frmIT_4B_G1.cmb24.Text
            Sheet_IT_4B_G1.Cells(6, 3).Value = frmIT_4B_G1.cmb25.Text
            Sheet_IT_4B_G1.Cells(7, 3).Value = frmIT_4B_G1.cmb26.Text
            Sheet_IT_4B_G1.Cells(8, 3).Value = frmIT_4B_G1.cmb27.Text
            Sheet_IT_4B_G1.Cells(9, 3).Value = frmIT_4B_G1.cmb28.Text
            Sheet_IT_4B_G1.Cells(10, 3).Value = frmIT_4B_G1.cmb29.Text
            Sheet_IT_4B_G1.Cells(11, 3).Value = frmIT_4B_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_4B_G1.Cells(2, 4).Value = frmIT_4B_G1.cmb41.Text
            Sheet_IT_4B_G1.Cells(3, 4).Value = frmIT_4B_G1.cmb42.Text
            Sheet_IT_4B_G1.Cells(4, 4).Value = frmIT_4B_G1.cmb43.Text
            Sheet_IT_4B_G1.Cells(5, 4).Value = frmIT_4B_G1.cmb44.Text
            Sheet_IT_4B_G1.Cells(6, 4).Value = frmIT_4B_G1.cmb45.Text
            Sheet_IT_4B_G1.Cells(7, 4).Value = frmIT_4B_G1.cmb46.Text
            Sheet_IT_4B_G1.Cells(8, 4).Value = frmIT_4B_G1.cmb47.Text
            Sheet_IT_4B_G1.Cells(9, 4).Value = frmIT_4B_G1.cmb48.Text
            Sheet_IT_4B_G1.Cells(10, 4).Value = frmIT_4B_G1.cmb49.Text
            Sheet_IT_4B_G1.Cells(11, 4).Value = frmIT_4B_G1.cmb50.Text

            ' Thursday

            Sheet_IT_4B_G1.Cells(2, 5).Value = frmIT_4B_G1.cmb61.Text
            Sheet_IT_4B_G1.Cells(3, 5).Value = frmIT_4B_G1.cmb62.Text
            Sheet_IT_4B_G1.Cells(4, 5).Value = frmIT_4B_G1.cmb63.Text
            Sheet_IT_4B_G1.Cells(5, 5).Value = frmIT_4B_G1.cmb64.Text
            Sheet_IT_4B_G1.Cells(6, 5).Value = frmIT_4B_G1.cmb65.Text
            Sheet_IT_4B_G1.Cells(7, 5).Value = frmIT_4B_G1.cmb66.Text
            Sheet_IT_4B_G1.Cells(8, 5).Value = frmIT_4B_G1.cmb67.Text
            Sheet_IT_4B_G1.Cells(9, 5).Value = frmIT_4B_G1.cmb68.Text
            Sheet_IT_4B_G1.Cells(10, 5).Value = frmIT_4B_G1.cmb69.Text
            Sheet_IT_4B_G1.Cells(11, 5).Value = frmIT_4B_G1.cmb70.Text

            ' Friday

            Sheet_IT_4B_G1.Cells(2, 6).Value = frmIT_4B_G1.cmb81.Text
            Sheet_IT_4B_G1.Cells(3, 6).Value = frmIT_4B_G1.cmb82.Text
            Sheet_IT_4B_G1.Cells(4, 6).Value = frmIT_4B_G1.cmb83.Text
            Sheet_IT_4B_G1.Cells(5, 6).Value = frmIT_4B_G1.cmb84.Text
            Sheet_IT_4B_G1.Cells(6, 6).Value = frmIT_4B_G1.cmb85.Text
            Sheet_IT_4B_G1.Cells(7, 6).Value = frmIT_4B_G1.cmb86.Text
            Sheet_IT_4B_G1.Cells(8, 6).Value = frmIT_4B_G1.cmb87.Text
            Sheet_IT_4B_G1.Cells(9, 6).Value = frmIT_4B_G1.cmb88.Text
            Sheet_IT_4B_G1.Cells(10, 6).Value = frmIT_4B_G1.cmb89.Text
            Sheet_IT_4B_G1.Cells(11, 6).Value = frmIT_4B_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4B_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4B_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4B_G2.Cells(1, 1).Value = ""
            Sheet_IT_4B_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4B_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4B_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4B_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4B_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4B_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4B_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4B_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4B_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4B_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4B_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_4B_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4B_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4B_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4B_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4B_G2.Cells(2, 2).Value = frmIT_4B_G2.cmb1.Text
            Sheet_IT_4B_G2.Cells(3, 2).Value = frmIT_4B_G2.cmb2.Text
            Sheet_IT_4B_G2.Cells(4, 2).Value = frmIT_4B_G2.cmb3.Text
            Sheet_IT_4B_G2.Cells(5, 2).Value = frmIT_4B_G2.cmb4.Text
            Sheet_IT_4B_G2.Cells(6, 2).Value = frmIT_4B_G2.cmb5.Text
            Sheet_IT_4B_G2.Cells(7, 2).Value = frmIT_4B_G2.cmb6.Text
            Sheet_IT_4B_G2.Cells(8, 2).Value = frmIT_4B_G2.cmb7.Text
            Sheet_IT_4B_G2.Cells(9, 2).Value = frmIT_4B_G2.cmb8.Text
            Sheet_IT_4B_G2.Cells(10, 2).Value = frmIT_4B_G2.cmb9.Text
            Sheet_IT_4B_G2.Cells(11, 2).Value = frmIT_4B_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_4B_G2.Cells(2, 3).Value = frmIT_4B_G2.cmb21.Text
            Sheet_IT_4B_G2.Cells(3, 3).Value = frmIT_4B_G2.cmb22.Text
            Sheet_IT_4B_G2.Cells(4, 3).Value = frmIT_4B_G2.cmb23.Text
            Sheet_IT_4B_G2.Cells(5, 3).Value = frmIT_4B_G2.cmb24.Text
            Sheet_IT_4B_G2.Cells(6, 3).Value = frmIT_4B_G2.cmb25.Text
            Sheet_IT_4B_G2.Cells(7, 3).Value = frmIT_4B_G2.cmb26.Text
            Sheet_IT_4B_G2.Cells(8, 3).Value = frmIT_4B_G2.cmb27.Text
            Sheet_IT_4B_G2.Cells(9, 3).Value = frmIT_4B_G2.cmb28.Text
            Sheet_IT_4B_G2.Cells(10, 3).Value = frmIT_4B_G2.cmb29.Text
            Sheet_IT_4B_G2.Cells(11, 3).Value = frmIT_4B_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_4B_G2.Cells(2, 4).Value = frmIT_4B_G2.cmb41.Text
            Sheet_IT_4B_G2.Cells(3, 4).Value = frmIT_4B_G2.cmb42.Text
            Sheet_IT_4B_G2.Cells(4, 4).Value = frmIT_4B_G2.cmb43.Text
            Sheet_IT_4B_G2.Cells(5, 4).Value = frmIT_4B_G2.cmb44.Text
            Sheet_IT_4B_G2.Cells(6, 4).Value = frmIT_4B_G2.cmb45.Text
            Sheet_IT_4B_G2.Cells(7, 4).Value = frmIT_4B_G2.cmb46.Text
            Sheet_IT_4B_G2.Cells(8, 4).Value = frmIT_4B_G2.cmb47.Text
            Sheet_IT_4B_G2.Cells(9, 4).Value = frmIT_4B_G2.cmb48.Text
            Sheet_IT_4B_G2.Cells(10, 4).Value = frmIT_4B_G2.cmb49.Text
            Sheet_IT_4B_G2.Cells(11, 4).Value = frmIT_4B_G2.cmb50.Text

            ' Thursday

            Sheet_IT_4B_G2.Cells(2, 5).Value = frmIT_4B_G2.cmb61.Text
            Sheet_IT_4B_G2.Cells(3, 5).Value = frmIT_4B_G2.cmb62.Text
            Sheet_IT_4B_G2.Cells(4, 5).Value = frmIT_4B_G2.cmb63.Text
            Sheet_IT_4B_G2.Cells(5, 5).Value = frmIT_4B_G2.cmb64.Text
            Sheet_IT_4B_G2.Cells(6, 5).Value = frmIT_4B_G2.cmb65.Text
            Sheet_IT_4B_G2.Cells(7, 5).Value = frmIT_4B_G2.cmb66.Text
            Sheet_IT_4B_G2.Cells(8, 5).Value = frmIT_4B_G2.cmb67.Text
            Sheet_IT_4B_G2.Cells(9, 5).Value = frmIT_4B_G2.cmb68.Text
            Sheet_IT_4B_G2.Cells(10, 5).Value = frmIT_4B_G2.cmb69.Text
            Sheet_IT_4B_G2.Cells(11, 5).Value = frmIT_4B_G2.cmb70.Text

            ' Friday

            Sheet_IT_4B_G2.Cells(2, 6).Value = frmIT_4B_G2.cmb81.Text
            Sheet_IT_4B_G2.Cells(3, 6).Value = frmIT_4B_G2.cmb82.Text
            Sheet_IT_4B_G2.Cells(4, 6).Value = frmIT_4B_G2.cmb83.Text
            Sheet_IT_4B_G2.Cells(5, 6).Value = frmIT_4B_G2.cmb84.Text
            Sheet_IT_4B_G2.Cells(6, 6).Value = frmIT_4B_G2.cmb85.Text
            Sheet_IT_4B_G2.Cells(7, 6).Value = frmIT_4B_G2.cmb86.Text
            Sheet_IT_4B_G2.Cells(8, 6).Value = frmIT_4B_G2.cmb87.Text
            Sheet_IT_4B_G2.Cells(9, 6).Value = frmIT_4B_G2.cmb88.Text
            Sheet_IT_4B_G2.Cells(10, 6).Value = frmIT_4B_G2.cmb89.Text
            Sheet_IT_4B_G2.Cells(11, 6).Value = frmIT_4B_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_4B_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_4B_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_4B_G3.Cells(1, 1).Value = ""
            Sheet_IT_4B_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_4B_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_4B_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_4B_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_4B_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_4B_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_4B_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_4B_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_4B_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_4B_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_4B_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_4B_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_4B_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_4B_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_4B_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_4B_G3.Cells(2, 2).Value = frmIT_4B_G3.cmb1.Text
            Sheet_IT_4B_G3.Cells(3, 2).Value = frmIT_4B_G3.cmb2.Text
            Sheet_IT_4B_G3.Cells(4, 2).Value = frmIT_4B_G3.cmb3.Text
            Sheet_IT_4B_G3.Cells(5, 2).Value = frmIT_4B_G3.cmb4.Text
            Sheet_IT_4B_G3.Cells(6, 2).Value = frmIT_4B_G3.cmb5.Text
            Sheet_IT_4B_G3.Cells(7, 2).Value = frmIT_4B_G3.cmb6.Text
            Sheet_IT_4B_G3.Cells(8, 2).Value = frmIT_4B_G3.cmb7.Text
            Sheet_IT_4B_G3.Cells(9, 2).Value = frmIT_4B_G3.cmb8.Text
            Sheet_IT_4B_G3.Cells(10, 2).Value = frmIT_4B_G3.cmb9.Text
            Sheet_IT_4B_G3.Cells(11, 2).Value = frmIT_4B_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_4B_G3.Cells(2, 3).Value = frmIT_4B_G3.cmb21.Text
            Sheet_IT_4B_G3.Cells(3, 3).Value = frmIT_4B_G3.cmb22.Text
            Sheet_IT_4B_G3.Cells(4, 3).Value = frmIT_4B_G3.cmb23.Text
            Sheet_IT_4B_G3.Cells(5, 3).Value = frmIT_4B_G3.cmb24.Text
            Sheet_IT_4B_G3.Cells(6, 3).Value = frmIT_4B_G3.cmb25.Text
            Sheet_IT_4B_G3.Cells(7, 3).Value = frmIT_4B_G3.cmb26.Text
            Sheet_IT_4B_G3.Cells(8, 3).Value = frmIT_4B_G3.cmb27.Text
            Sheet_IT_4B_G3.Cells(9, 3).Value = frmIT_4B_G3.cmb28.Text
            Sheet_IT_4B_G3.Cells(10, 3).Value = frmIT_4B_G3.cmb29.Text
            Sheet_IT_4B_G3.Cells(11, 3).Value = frmIT_4B_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_4B_G3.Cells(2, 4).Value = frmIT_4B_G3.cmb41.Text
            Sheet_IT_4B_G3.Cells(3, 4).Value = frmIT_4B_G3.cmb42.Text
            Sheet_IT_4B_G3.Cells(4, 4).Value = frmIT_4B_G3.cmb43.Text
            Sheet_IT_4B_G3.Cells(5, 4).Value = frmIT_4B_G3.cmb44.Text
            Sheet_IT_4B_G3.Cells(6, 4).Value = frmIT_4B_G3.cmb45.Text
            Sheet_IT_4B_G3.Cells(7, 4).Value = frmIT_4B_G3.cmb46.Text
            Sheet_IT_4B_G3.Cells(8, 4).Value = frmIT_4B_G3.cmb47.Text
            Sheet_IT_4B_G3.Cells(9, 4).Value = frmIT_4B_G3.cmb48.Text
            Sheet_IT_4B_G3.Cells(10, 4).Value = frmIT_4B_G3.cmb49.Text
            Sheet_IT_4B_G3.Cells(11, 4).Value = frmIT_4B_G3.cmb50.Text

            ' Thursday

            Sheet_IT_4B_G3.Cells(2, 5).Value = frmIT_4B_G3.cmb61.Text
            Sheet_IT_4B_G3.Cells(3, 5).Value = frmIT_4B_G3.cmb62.Text
            Sheet_IT_4B_G3.Cells(4, 5).Value = frmIT_4B_G3.cmb63.Text
            Sheet_IT_4B_G3.Cells(5, 5).Value = frmIT_4B_G3.cmb64.Text
            Sheet_IT_4B_G3.Cells(6, 5).Value = frmIT_4B_G3.cmb65.Text
            Sheet_IT_4B_G3.Cells(7, 5).Value = frmIT_4B_G3.cmb66.Text
            Sheet_IT_4B_G3.Cells(8, 5).Value = frmIT_4B_G3.cmb67.Text
            Sheet_IT_4B_G3.Cells(9, 5).Value = frmIT_4B_G3.cmb68.Text
            Sheet_IT_4B_G3.Cells(10, 5).Value = frmIT_4B_G3.cmb69.Text
            Sheet_IT_4B_G3.Cells(11, 5).Value = frmIT_4B_G3.cmb70.Text

            ' Friday

            Sheet_IT_4B_G3.Cells(2, 6).Value = frmIT_4B_G3.cmb81.Text
            Sheet_IT_4B_G3.Cells(3, 6).Value = frmIT_4B_G3.cmb82.Text
            Sheet_IT_4B_G3.Cells(4, 6).Value = frmIT_4B_G3.cmb83.Text
            Sheet_IT_4B_G3.Cells(5, 6).Value = frmIT_4B_G3.cmb84.Text
            Sheet_IT_4B_G3.Cells(6, 6).Value = frmIT_4B_G3.cmb85.Text
            Sheet_IT_4B_G3.Cells(7, 6).Value = frmIT_4B_G3.cmb86.Text
            Sheet_IT_4B_G3.Cells(8, 6).Value = frmIT_4B_G3.cmb87.Text
            Sheet_IT_4B_G3.Cells(9, 6).Value = frmIT_4B_G3.cmb88.Text
            Sheet_IT_4B_G3.Cells(10, 6).Value = frmIT_4B_G3.cmb89.Text
            Sheet_IT_4B_G3.Cells(11, 6).Value = frmIT_4B_G3.cmb90.Text
            '
            '
            '
            '
            '
            '
            '
            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_8A_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_8A_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_8A_G1.Cells(1, 1).Value = ""
            Sheet_IT_8A_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_8A_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_8A_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_8A_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_8A_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_8A_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_8A_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_8A_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_8A_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_8A_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_8A_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_8A_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_8A_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_8A_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_8A_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_8A_G1.Cells(2, 2).Value = frmIT_8A_G1.cmb1.Text
            Sheet_IT_8A_G1.Cells(3, 2).Value = frmIT_8A_G1.cmb2.Text
            Sheet_IT_8A_G1.Cells(4, 2).Value = frmIT_8A_G1.cmb3.Text
            Sheet_IT_8A_G1.Cells(5, 2).Value = frmIT_8A_G1.cmb4.Text
            Sheet_IT_8A_G1.Cells(6, 2).Value = frmIT_8A_G1.cmb5.Text
            Sheet_IT_8A_G1.Cells(7, 2).Value = frmIT_8A_G1.cmb6.Text
            Sheet_IT_8A_G1.Cells(8, 2).Value = frmIT_8A_G1.cmb7.Text
            Sheet_IT_8A_G1.Cells(9, 2).Value = frmIT_8A_G1.cmb8.Text
            Sheet_IT_8A_G1.Cells(10, 2).Value = frmIT_8A_G1.cmb9.Text
            Sheet_IT_8A_G1.Cells(11, 2).Value = frmIT_8A_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_8A_G1.Cells(2, 3).Value = frmIT_8A_G1.cmb21.Text
            Sheet_IT_8A_G1.Cells(3, 3).Value = frmIT_8A_G1.cmb22.Text
            Sheet_IT_8A_G1.Cells(4, 3).Value = frmIT_8A_G1.cmb23.Text
            Sheet_IT_8A_G1.Cells(5, 3).Value = frmIT_8A_G1.cmb24.Text
            Sheet_IT_8A_G1.Cells(6, 3).Value = frmIT_8A_G1.cmb25.Text
            Sheet_IT_8A_G1.Cells(7, 3).Value = frmIT_8A_G1.cmb26.Text
            Sheet_IT_8A_G1.Cells(8, 3).Value = frmIT_8A_G1.cmb27.Text
            Sheet_IT_8A_G1.Cells(9, 3).Value = frmIT_8A_G1.cmb28.Text
            Sheet_IT_8A_G1.Cells(10, 3).Value = frmIT_8A_G1.cmb29.Text
            Sheet_IT_8A_G1.Cells(11, 3).Value = frmIT_8A_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_8A_G1.Cells(2, 4).Value = frmIT_8A_G1.cmb41.Text
            Sheet_IT_8A_G1.Cells(3, 4).Value = frmIT_8A_G1.cmb42.Text
            Sheet_IT_8A_G1.Cells(4, 4).Value = frmIT_8A_G1.cmb43.Text
            Sheet_IT_8A_G1.Cells(5, 4).Value = frmIT_8A_G1.cmb44.Text
            Sheet_IT_8A_G1.Cells(6, 4).Value = frmIT_8A_G1.cmb45.Text
            Sheet_IT_8A_G1.Cells(7, 4).Value = frmIT_8A_G1.cmb46.Text
            Sheet_IT_8A_G1.Cells(8, 4).Value = frmIT_8A_G1.cmb47.Text
            Sheet_IT_8A_G1.Cells(9, 4).Value = frmIT_8A_G1.cmb48.Text
            Sheet_IT_8A_G1.Cells(10, 4).Value = frmIT_8A_G1.cmb49.Text
            Sheet_IT_8A_G1.Cells(11, 4).Value = frmIT_8A_G1.cmb50.Text

            ' Thursday

            Sheet_IT_8A_G1.Cells(2, 5).Value = frmIT_8A_G1.cmb61.Text
            Sheet_IT_8A_G1.Cells(3, 5).Value = frmIT_8A_G1.cmb62.Text
            Sheet_IT_8A_G1.Cells(4, 5).Value = frmIT_8A_G1.cmb63.Text
            Sheet_IT_8A_G1.Cells(5, 5).Value = frmIT_8A_G1.cmb64.Text
            Sheet_IT_8A_G1.Cells(6, 5).Value = frmIT_8A_G1.cmb65.Text
            Sheet_IT_8A_G1.Cells(7, 5).Value = frmIT_8A_G1.cmb66.Text
            Sheet_IT_8A_G1.Cells(8, 5).Value = frmIT_8A_G1.cmb67.Text
            Sheet_IT_8A_G1.Cells(9, 5).Value = frmIT_8A_G1.cmb68.Text
            Sheet_IT_8A_G1.Cells(10, 5).Value = frmIT_8A_G1.cmb69.Text
            Sheet_IT_8A_G1.Cells(11, 5).Value = frmIT_8A_G1.cmb70.Text

            ' Friday

            Sheet_IT_8A_G1.Cells(2, 6).Value = frmIT_8A_G1.cmb81.Text
            Sheet_IT_8A_G1.Cells(3, 6).Value = frmIT_8A_G1.cmb82.Text
            Sheet_IT_8A_G1.Cells(4, 6).Value = frmIT_8A_G1.cmb83.Text
            Sheet_IT_8A_G1.Cells(5, 6).Value = frmIT_8A_G1.cmb84.Text
            Sheet_IT_8A_G1.Cells(6, 6).Value = frmIT_8A_G1.cmb85.Text
            Sheet_IT_8A_G1.Cells(7, 6).Value = frmIT_8A_G1.cmb86.Text
            Sheet_IT_8A_G1.Cells(8, 6).Value = frmIT_8A_G1.cmb87.Text
            Sheet_IT_8A_G1.Cells(9, 6).Value = frmIT_8A_G1.cmb88.Text
            Sheet_IT_8A_G1.Cells(10, 6).Value = frmIT_8A_G1.cmb89.Text
            Sheet_IT_8A_G1.Cells(11, 6).Value = frmIT_8A_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_8A_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_8A_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_8A_G2.Cells(1, 1).Value = ""
            Sheet_IT_8A_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_8A_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_8A_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_8A_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_8A_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_8A_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_8A_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_8A_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_8A_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_8A_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_8A_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_8A_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_8A_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_8A_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_8A_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_8A_G2.Cells(2, 2).Value = frmIT_8A_G2.cmb1.Text
            Sheet_IT_8A_G2.Cells(3, 2).Value = frmIT_8A_G2.cmb2.Text
            Sheet_IT_8A_G2.Cells(4, 2).Value = frmIT_8A_G2.cmb3.Text
            Sheet_IT_8A_G2.Cells(5, 2).Value = frmIT_8A_G2.cmb4.Text
            Sheet_IT_8A_G2.Cells(6, 2).Value = frmIT_8A_G2.cmb5.Text
            Sheet_IT_8A_G2.Cells(7, 2).Value = frmIT_8A_G2.cmb6.Text
            Sheet_IT_8A_G2.Cells(8, 2).Value = frmIT_8A_G2.cmb7.Text
            Sheet_IT_8A_G2.Cells(9, 2).Value = frmIT_8A_G2.cmb8.Text
            Sheet_IT_8A_G2.Cells(10, 2).Value = frmIT_8A_G2.cmb9.Text
            Sheet_IT_8A_G2.Cells(11, 2).Value = frmIT_8A_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_8A_G2.Cells(2, 3).Value = frmIT_8A_G2.cmb21.Text
            Sheet_IT_8A_G2.Cells(3, 3).Value = frmIT_8A_G2.cmb22.Text
            Sheet_IT_8A_G2.Cells(4, 3).Value = frmIT_8A_G2.cmb23.Text
            Sheet_IT_8A_G2.Cells(5, 3).Value = frmIT_8A_G2.cmb24.Text
            Sheet_IT_8A_G2.Cells(6, 3).Value = frmIT_8A_G2.cmb25.Text
            Sheet_IT_8A_G2.Cells(7, 3).Value = frmIT_8A_G2.cmb26.Text
            Sheet_IT_8A_G2.Cells(8, 3).Value = frmIT_8A_G2.cmb27.Text
            Sheet_IT_8A_G2.Cells(9, 3).Value = frmIT_8A_G2.cmb28.Text
            Sheet_IT_8A_G2.Cells(10, 3).Value = frmIT_8A_G2.cmb29.Text
            Sheet_IT_8A_G2.Cells(11, 3).Value = frmIT_8A_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_8A_G2.Cells(2, 4).Value = frmIT_8A_G2.cmb41.Text
            Sheet_IT_8A_G2.Cells(3, 4).Value = frmIT_8A_G2.cmb42.Text
            Sheet_IT_8A_G2.Cells(4, 4).Value = frmIT_8A_G2.cmb43.Text
            Sheet_IT_8A_G2.Cells(5, 4).Value = frmIT_8A_G2.cmb44.Text
            Sheet_IT_8A_G2.Cells(6, 4).Value = frmIT_8A_G2.cmb45.Text
            Sheet_IT_8A_G2.Cells(7, 4).Value = frmIT_8A_G2.cmb46.Text
            Sheet_IT_8A_G2.Cells(8, 4).Value = frmIT_8A_G2.cmb47.Text
            Sheet_IT_8A_G2.Cells(9, 4).Value = frmIT_8A_G2.cmb48.Text
            Sheet_IT_8A_G2.Cells(10, 4).Value = frmIT_8A_G2.cmb49.Text
            Sheet_IT_8A_G2.Cells(11, 4).Value = frmIT_8A_G2.cmb50.Text

            ' Thursday

            Sheet_IT_8A_G2.Cells(2, 5).Value = frmIT_8A_G2.cmb61.Text
            Sheet_IT_8A_G2.Cells(3, 5).Value = frmIT_8A_G2.cmb62.Text
            Sheet_IT_8A_G2.Cells(4, 5).Value = frmIT_8A_G2.cmb63.Text
            Sheet_IT_8A_G2.Cells(5, 5).Value = frmIT_8A_G2.cmb64.Text
            Sheet_IT_8A_G2.Cells(6, 5).Value = frmIT_8A_G2.cmb65.Text
            Sheet_IT_8A_G2.Cells(7, 5).Value = frmIT_8A_G2.cmb66.Text
            Sheet_IT_8A_G2.Cells(8, 5).Value = frmIT_8A_G2.cmb67.Text
            Sheet_IT_8A_G2.Cells(9, 5).Value = frmIT_8A_G2.cmb68.Text
            Sheet_IT_8A_G2.Cells(10, 5).Value = frmIT_8A_G2.cmb69.Text
            Sheet_IT_8A_G2.Cells(11, 5).Value = frmIT_8A_G2.cmb70.Text

            ' Friday

            Sheet_IT_8A_G2.Cells(2, 6).Value = frmIT_8A_G2.cmb81.Text
            Sheet_IT_8A_G2.Cells(3, 6).Value = frmIT_8A_G2.cmb82.Text
            Sheet_IT_8A_G2.Cells(4, 6).Value = frmIT_8A_G2.cmb83.Text
            Sheet_IT_8A_G2.Cells(5, 6).Value = frmIT_8A_G2.cmb84.Text
            Sheet_IT_8A_G2.Cells(6, 6).Value = frmIT_8A_G2.cmb85.Text
            Sheet_IT_8A_G2.Cells(7, 6).Value = frmIT_8A_G2.cmb86.Text
            Sheet_IT_8A_G2.Cells(8, 6).Value = frmIT_8A_G2.cmb87.Text
            Sheet_IT_8A_G2.Cells(9, 6).Value = frmIT_8A_G2.cmb88.Text
            Sheet_IT_8A_G2.Cells(10, 6).Value = frmIT_8A_G2.cmb89.Text
            Sheet_IT_8A_G2.Cells(11, 6).Value = frmIT_8A_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_8A_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_8A_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_8A_G3.Cells(1, 1).Value = ""
            Sheet_IT_8A_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_8A_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_8A_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_8A_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_8A_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_8A_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_8A_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_8A_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_8A_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_8A_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_8A_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_8A_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_8A_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_8A_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_8A_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_8A_G3.Cells(2, 2).Value = frmIT_8A_G3.cmb1.Text
            Sheet_IT_8A_G3.Cells(3, 2).Value = frmIT_8A_G3.cmb2.Text
            Sheet_IT_8A_G3.Cells(4, 2).Value = frmIT_8A_G3.cmb3.Text
            Sheet_IT_8A_G3.Cells(5, 2).Value = frmIT_8A_G3.cmb4.Text
            Sheet_IT_8A_G3.Cells(6, 2).Value = frmIT_8A_G3.cmb5.Text
            Sheet_IT_8A_G3.Cells(7, 2).Value = frmIT_8A_G3.cmb6.Text
            Sheet_IT_8A_G3.Cells(8, 2).Value = frmIT_8A_G3.cmb7.Text
            Sheet_IT_8A_G3.Cells(9, 2).Value = frmIT_8A_G3.cmb8.Text
            Sheet_IT_8A_G3.Cells(10, 2).Value = frmIT_8A_G3.cmb9.Text
            Sheet_IT_8A_G3.Cells(11, 2).Value = frmIT_8A_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_8A_G3.Cells(2, 3).Value = frmIT_8A_G3.cmb21.Text
            Sheet_IT_8A_G3.Cells(3, 3).Value = frmIT_8A_G3.cmb22.Text
            Sheet_IT_8A_G3.Cells(4, 3).Value = frmIT_8A_G3.cmb23.Text
            Sheet_IT_8A_G3.Cells(5, 3).Value = frmIT_8A_G3.cmb24.Text
            Sheet_IT_8A_G3.Cells(6, 3).Value = frmIT_8A_G3.cmb25.Text
            Sheet_IT_8A_G3.Cells(7, 3).Value = frmIT_8A_G3.cmb26.Text
            Sheet_IT_8A_G3.Cells(8, 3).Value = frmIT_8A_G3.cmb27.Text
            Sheet_IT_8A_G3.Cells(9, 3).Value = frmIT_8A_G3.cmb28.Text
            Sheet_IT_8A_G3.Cells(10, 3).Value = frmIT_8A_G3.cmb29.Text
            Sheet_IT_8A_G3.Cells(11, 3).Value = frmIT_8A_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_8A_G3.Cells(2, 4).Value = frmIT_8A_G3.cmb41.Text
            Sheet_IT_8A_G3.Cells(3, 4).Value = frmIT_8A_G3.cmb42.Text
            Sheet_IT_8A_G3.Cells(4, 4).Value = frmIT_8A_G3.cmb43.Text
            Sheet_IT_8A_G3.Cells(5, 4).Value = frmIT_8A_G3.cmb44.Text
            Sheet_IT_8A_G3.Cells(6, 4).Value = frmIT_8A_G3.cmb45.Text
            Sheet_IT_8A_G3.Cells(7, 4).Value = frmIT_8A_G3.cmb46.Text
            Sheet_IT_8A_G3.Cells(8, 4).Value = frmIT_8A_G3.cmb47.Text
            Sheet_IT_8A_G3.Cells(9, 4).Value = frmIT_8A_G3.cmb48.Text
            Sheet_IT_8A_G3.Cells(10, 4).Value = frmIT_8A_G3.cmb49.Text
            Sheet_IT_8A_G3.Cells(11, 4).Value = frmIT_8A_G3.cmb50.Text

            ' Thursday

            Sheet_IT_8A_G3.Cells(2, 5).Value = frmIT_8A_G3.cmb61.Text
            Sheet_IT_8A_G3.Cells(3, 5).Value = frmIT_8A_G3.cmb62.Text
            Sheet_IT_8A_G3.Cells(4, 5).Value = frmIT_8A_G3.cmb63.Text
            Sheet_IT_8A_G3.Cells(5, 5).Value = frmIT_8A_G3.cmb64.Text
            Sheet_IT_8A_G3.Cells(6, 5).Value = frmIT_8A_G3.cmb65.Text
            Sheet_IT_8A_G3.Cells(7, 5).Value = frmIT_8A_G3.cmb66.Text
            Sheet_IT_8A_G3.Cells(8, 5).Value = frmIT_8A_G3.cmb67.Text
            Sheet_IT_8A_G3.Cells(9, 5).Value = frmIT_8A_G3.cmb68.Text
            Sheet_IT_8A_G3.Cells(10, 5).Value = frmIT_8A_G3.cmb69.Text
            Sheet_IT_8A_G3.Cells(11, 5).Value = frmIT_8A_G3.cmb70.Text

            ' Friday

            Sheet_IT_8A_G3.Cells(2, 6).Value = frmIT_8A_G3.cmb81.Text
            Sheet_IT_8A_G3.Cells(3, 6).Value = frmIT_8A_G3.cmb82.Text
            Sheet_IT_8A_G3.Cells(4, 6).Value = frmIT_8A_G3.cmb83.Text
            Sheet_IT_8A_G3.Cells(5, 6).Value = frmIT_8A_G3.cmb84.Text
            Sheet_IT_8A_G3.Cells(6, 6).Value = frmIT_8A_G3.cmb85.Text
            Sheet_IT_8A_G3.Cells(7, 6).Value = frmIT_8A_G3.cmb86.Text
            Sheet_IT_8A_G3.Cells(8, 6).Value = frmIT_8A_G3.cmb87.Text
            Sheet_IT_8A_G3.Cells(9, 6).Value = frmIT_8A_G3.cmb88.Text
            Sheet_IT_8A_G3.Cells(10, 6).Value = frmIT_8A_G3.cmb89.Text
            Sheet_IT_8A_G3.Cells(11, 6).Value = frmIT_8A_G3.cmb90.Text
            ''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''
            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_8B_G1.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_8B_G1.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_8B_G1.Cells(1, 1).Value = ""
            Sheet_IT_8B_G1.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_8B_G1.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_8B_G1.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_8B_G1.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_8B_G1.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_8B_G1.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_8B_G1.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_8B_G1.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_8B_G1.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_8B_G1.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_8B_G1.Cells(2, 1).Value = "Monday"
            Sheet_IT_8B_G1.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_8B_G1.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_8B_G1.Cells(5, 1).Value = "Thursday"
            Sheet_IT_8B_G1.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_8B_G1.Cells(2, 2).Value = frmIT_8B_G1.cmb1.Text
            Sheet_IT_8B_G1.Cells(3, 2).Value = frmIT_8B_G1.cmb2.Text
            Sheet_IT_8B_G1.Cells(4, 2).Value = frmIT_8B_G1.cmb3.Text
            Sheet_IT_8B_G1.Cells(5, 2).Value = frmIT_8B_G1.cmb4.Text
            Sheet_IT_8B_G1.Cells(6, 2).Value = frmIT_8B_G1.cmb5.Text
            Sheet_IT_8B_G1.Cells(7, 2).Value = frmIT_8B_G1.cmb6.Text
            Sheet_IT_8B_G1.Cells(8, 2).Value = frmIT_8B_G1.cmb7.Text
            Sheet_IT_8B_G1.Cells(9, 2).Value = frmIT_8B_G1.cmb8.Text
            Sheet_IT_8B_G1.Cells(10, 2).Value = frmIT_8B_G1.cmb9.Text
            Sheet_IT_8B_G1.Cells(11, 2).Value = frmIT_8B_G1.cmb10.Text

            ' Tuesday

            Sheet_IT_8B_G1.Cells(2, 3).Value = frmIT_8B_G1.cmb21.Text
            Sheet_IT_8B_G1.Cells(3, 3).Value = frmIT_8B_G1.cmb22.Text
            Sheet_IT_8B_G1.Cells(4, 3).Value = frmIT_8B_G1.cmb23.Text
            Sheet_IT_8B_G1.Cells(5, 3).Value = frmIT_8B_G1.cmb24.Text
            Sheet_IT_8B_G1.Cells(6, 3).Value = frmIT_8B_G1.cmb25.Text
            Sheet_IT_8B_G1.Cells(7, 3).Value = frmIT_8B_G1.cmb26.Text
            Sheet_IT_8B_G1.Cells(8, 3).Value = frmIT_8B_G1.cmb27.Text
            Sheet_IT_8B_G1.Cells(9, 3).Value = frmIT_8B_G1.cmb28.Text
            Sheet_IT_8B_G1.Cells(10, 3).Value = frmIT_8B_G1.cmb29.Text
            Sheet_IT_8B_G1.Cells(11, 3).Value = frmIT_8B_G1.cmb30.Text

            ' Wednessday

            Sheet_IT_8B_G1.Cells(2, 4).Value = frmIT_8B_G1.cmb41.Text
            Sheet_IT_8B_G1.Cells(3, 4).Value = frmIT_8B_G1.cmb42.Text
            Sheet_IT_8B_G1.Cells(4, 4).Value = frmIT_8B_G1.cmb43.Text
            Sheet_IT_8B_G1.Cells(5, 4).Value = frmIT_8B_G1.cmb44.Text
            Sheet_IT_8B_G1.Cells(6, 4).Value = frmIT_8B_G1.cmb45.Text
            Sheet_IT_8B_G1.Cells(7, 4).Value = frmIT_8B_G1.cmb46.Text
            Sheet_IT_8B_G1.Cells(8, 4).Value = frmIT_8B_G1.cmb47.Text
            Sheet_IT_8B_G1.Cells(9, 4).Value = frmIT_8B_G1.cmb48.Text
            Sheet_IT_8B_G1.Cells(10, 4).Value = frmIT_8B_G1.cmb49.Text
            Sheet_IT_8B_G1.Cells(11, 4).Value = frmIT_8B_G1.cmb50.Text

            ' Thursday

            Sheet_IT_8B_G1.Cells(2, 5).Value = frmIT_8B_G1.cmb61.Text
            Sheet_IT_8B_G1.Cells(3, 5).Value = frmIT_8B_G1.cmb62.Text
            Sheet_IT_8B_G1.Cells(4, 5).Value = frmIT_8B_G1.cmb63.Text
            Sheet_IT_8B_G1.Cells(5, 5).Value = frmIT_8B_G1.cmb64.Text
            Sheet_IT_8B_G1.Cells(6, 5).Value = frmIT_8B_G1.cmb65.Text
            Sheet_IT_8B_G1.Cells(7, 5).Value = frmIT_8B_G1.cmb66.Text
            Sheet_IT_8B_G1.Cells(8, 5).Value = frmIT_8B_G1.cmb67.Text
            Sheet_IT_8B_G1.Cells(9, 5).Value = frmIT_8B_G1.cmb68.Text
            Sheet_IT_8B_G1.Cells(10, 5).Value = frmIT_8B_G1.cmb69.Text
            Sheet_IT_8B_G1.Cells(11, 5).Value = frmIT_8B_G1.cmb70.Text

            ' Friday

            Sheet_IT_8B_G1.Cells(2, 6).Value = frmIT_8B_G1.cmb81.Text
            Sheet_IT_8B_G1.Cells(3, 6).Value = frmIT_8B_G1.cmb82.Text
            Sheet_IT_8B_G1.Cells(4, 6).Value = frmIT_8B_G1.cmb83.Text
            Sheet_IT_8B_G1.Cells(5, 6).Value = frmIT_8B_G1.cmb84.Text
            Sheet_IT_8B_G1.Cells(6, 6).Value = frmIT_8B_G1.cmb85.Text
            Sheet_IT_8B_G1.Cells(7, 6).Value = frmIT_8B_G1.cmb86.Text
            Sheet_IT_8B_G1.Cells(8, 6).Value = frmIT_8B_G1.cmb87.Text
            Sheet_IT_8B_G1.Cells(9, 6).Value = frmIT_8B_G1.cmb88.Text
            Sheet_IT_8B_G1.Cells(10, 6).Value = frmIT_8B_G1.cmb89.Text
            Sheet_IT_8B_G1.Cells(11, 6).Value = frmIT_8B_G1.cmb90.Text

            '''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_8B_G2.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_8B_G2.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_8B_G2.Cells(1, 1).Value = ""
            Sheet_IT_8B_G2.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_8B_G2.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_8B_G2.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_8B_G2.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_8B_G2.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_8B_G2.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_8B_G2.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_8B_G2.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_8B_G2.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_8B_G2.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_8B_G2.Cells(2, 1).Value = "Monday"
            Sheet_IT_8B_G2.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_8B_G2.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_8B_G2.Cells(5, 1).Value = "Thursday"
            Sheet_IT_8B_G2.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_8B_G2.Cells(2, 2).Value = frmIT_8B_G2.cmb1.Text
            Sheet_IT_8B_G2.Cells(3, 2).Value = frmIT_8B_G2.cmb2.Text
            Sheet_IT_8B_G2.Cells(4, 2).Value = frmIT_8B_G2.cmb3.Text
            Sheet_IT_8B_G2.Cells(5, 2).Value = frmIT_8B_G2.cmb4.Text
            Sheet_IT_8B_G2.Cells(6, 2).Value = frmIT_8B_G2.cmb5.Text
            Sheet_IT_8B_G2.Cells(7, 2).Value = frmIT_8B_G2.cmb6.Text
            Sheet_IT_8B_G2.Cells(8, 2).Value = frmIT_8B_G2.cmb7.Text
            Sheet_IT_8B_G2.Cells(9, 2).Value = frmIT_8B_G2.cmb8.Text
            Sheet_IT_8B_G2.Cells(10, 2).Value = frmIT_8B_G2.cmb9.Text
            Sheet_IT_8B_G2.Cells(11, 2).Value = frmIT_8B_G2.cmb10.Text

            ' Tuesday

            Sheet_IT_8B_G2.Cells(2, 3).Value = frmIT_8B_G2.cmb21.Text
            Sheet_IT_8B_G2.Cells(3, 3).Value = frmIT_8B_G2.cmb22.Text
            Sheet_IT_8B_G2.Cells(4, 3).Value = frmIT_8B_G2.cmb23.Text
            Sheet_IT_8B_G2.Cells(5, 3).Value = frmIT_8B_G2.cmb24.Text
            Sheet_IT_8B_G2.Cells(6, 3).Value = frmIT_8B_G2.cmb25.Text
            Sheet_IT_8B_G2.Cells(7, 3).Value = frmIT_8B_G2.cmb26.Text
            Sheet_IT_8B_G2.Cells(8, 3).Value = frmIT_8B_G2.cmb27.Text
            Sheet_IT_8B_G2.Cells(9, 3).Value = frmIT_8B_G2.cmb28.Text
            Sheet_IT_8B_G2.Cells(10, 3).Value = frmIT_8B_G2.cmb29.Text
            Sheet_IT_8B_G2.Cells(11, 3).Value = frmIT_8B_G2.cmb30.Text

            ' Wednessday

            Sheet_IT_8B_G2.Cells(2, 4).Value = frmIT_8B_G2.cmb41.Text
            Sheet_IT_8B_G2.Cells(3, 4).Value = frmIT_8B_G2.cmb42.Text
            Sheet_IT_8B_G2.Cells(4, 4).Value = frmIT_8B_G2.cmb43.Text
            Sheet_IT_8B_G2.Cells(5, 4).Value = frmIT_8B_G2.cmb44.Text
            Sheet_IT_8B_G2.Cells(6, 4).Value = frmIT_8B_G2.cmb45.Text
            Sheet_IT_8B_G2.Cells(7, 4).Value = frmIT_8B_G2.cmb46.Text
            Sheet_IT_8B_G2.Cells(8, 4).Value = frmIT_8B_G2.cmb47.Text
            Sheet_IT_8B_G2.Cells(9, 4).Value = frmIT_8B_G2.cmb48.Text
            Sheet_IT_8B_G2.Cells(10, 4).Value = frmIT_8B_G2.cmb49.Text
            Sheet_IT_8B_G2.Cells(11, 4).Value = frmIT_8B_G2.cmb50.Text

            ' Thursday

            Sheet_IT_8B_G2.Cells(2, 5).Value = frmIT_8B_G2.cmb61.Text
            Sheet_IT_8B_G2.Cells(3, 5).Value = frmIT_8B_G2.cmb62.Text
            Sheet_IT_8B_G2.Cells(4, 5).Value = frmIT_8B_G2.cmb63.Text
            Sheet_IT_8B_G2.Cells(5, 5).Value = frmIT_8B_G2.cmb64.Text
            Sheet_IT_8B_G2.Cells(6, 5).Value = frmIT_8B_G2.cmb65.Text
            Sheet_IT_8B_G2.Cells(7, 5).Value = frmIT_8B_G2.cmb66.Text
            Sheet_IT_8B_G2.Cells(8, 5).Value = frmIT_8B_G2.cmb67.Text
            Sheet_IT_8B_G2.Cells(9, 5).Value = frmIT_8B_G2.cmb68.Text
            Sheet_IT_8B_G2.Cells(10, 5).Value = frmIT_8B_G2.cmb69.Text
            Sheet_IT_8B_G2.Cells(11, 5).Value = frmIT_8B_G2.cmb70.Text

            ' Friday

            Sheet_IT_8B_G2.Cells(2, 6).Value = frmIT_8B_G2.cmb81.Text
            Sheet_IT_8B_G2.Cells(3, 6).Value = frmIT_8B_G2.cmb82.Text
            Sheet_IT_8B_G2.Cells(4, 6).Value = frmIT_8B_G2.cmb83.Text
            Sheet_IT_8B_G2.Cells(5, 6).Value = frmIT_8B_G2.cmb84.Text
            Sheet_IT_8B_G2.Cells(6, 6).Value = frmIT_8B_G2.cmb85.Text
            Sheet_IT_8B_G2.Cells(7, 6).Value = frmIT_8B_G2.cmb86.Text
            Sheet_IT_8B_G2.Cells(8, 6).Value = frmIT_8B_G2.cmb87.Text
            Sheet_IT_8B_G2.Cells(9, 6).Value = frmIT_8B_G2.cmb88.Text
            Sheet_IT_8B_G2.Cells(10, 6).Value = frmIT_8B_G2.cmb89.Text
            Sheet_IT_8B_G2.Cells(11, 6).Value = frmIT_8B_G2.cmb90.Text

            '''''''''''''''''''''''''''''''''''''''''''

            ' Format A1:K1 as bold, vertical alignment = center.
            With Sheet_IT_8B_G3.Range("A1", "K1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            ' Format A1:A6 as bold, vertical alignment = center.
            With Sheet_IT_8B_G3.Range("A1", "A6")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Add table headers going cell by cell.
            Sheet_IT_8B_G3.Cells(1, 1).Value = ""
            Sheet_IT_8B_G3.Cells(1, 2).Value = "8 - 9 AM"
            Sheet_IT_8B_G3.Cells(1, 3).Value = "9 - 10 AM"
            Sheet_IT_8B_G3.Cells(1, 4).Value = "10 - 11 AM"
            Sheet_IT_8B_G3.Cells(1, 5).Value = "11 - 12 AM"
            Sheet_IT_8B_G3.Cells(1, 6).Value = "12 - 1 PM"
            Sheet_IT_8B_G3.Cells(1, 7).Value = "1 - 2 PM"
            Sheet_IT_8B_G3.Cells(1, 8).Value = "2 - 3 PM"
            Sheet_IT_8B_G3.Cells(1, 9).Value = "3 - 4 PM"
            Sheet_IT_8B_G3.Cells(1, 10).Value = "4 - 5 PM"
            Sheet_IT_8B_G3.Cells(1, 11).Value = "5 - 6 PM"
            Sheet_IT_8B_G3.Cells(2, 1).Value = "Monday"
            Sheet_IT_8B_G3.Cells(3, 1).Value = "Tuesday"
            Sheet_IT_8B_G3.Cells(4, 1).Value = "Wednessday"
            Sheet_IT_8B_G3.Cells(5, 1).Value = "Thursday"
            Sheet_IT_8B_G3.Cells(6, 1).Value = "Friday"

            ' Enter Values from ComboBox

            ' Monday

            Sheet_IT_8B_G3.Cells(2, 2).Value = frmIT_8B_G3.cmb1.Text
            Sheet_IT_8B_G3.Cells(3, 2).Value = frmIT_8B_G3.cmb2.Text
            Sheet_IT_8B_G3.Cells(4, 2).Value = frmIT_8B_G3.cmb3.Text
            Sheet_IT_8B_G3.Cells(5, 2).Value = frmIT_8B_G3.cmb4.Text
            Sheet_IT_8B_G3.Cells(6, 2).Value = frmIT_8B_G3.cmb5.Text
            Sheet_IT_8B_G3.Cells(7, 2).Value = frmIT_8B_G3.cmb6.Text
            Sheet_IT_8B_G3.Cells(8, 2).Value = frmIT_8B_G3.cmb7.Text
            Sheet_IT_8B_G3.Cells(9, 2).Value = frmIT_8B_G3.cmb8.Text
            Sheet_IT_8B_G3.Cells(10, 2).Value = frmIT_8B_G3.cmb9.Text
            Sheet_IT_8B_G3.Cells(11, 2).Value = frmIT_8B_G3.cmb10.Text

            ' Tuesday

            Sheet_IT_8B_G3.Cells(2, 3).Value = frmIT_8B_G3.cmb21.Text
            Sheet_IT_8B_G3.Cells(3, 3).Value = frmIT_8B_G3.cmb22.Text
            Sheet_IT_8B_G3.Cells(4, 3).Value = frmIT_8B_G3.cmb23.Text
            Sheet_IT_8B_G3.Cells(5, 3).Value = frmIT_8B_G3.cmb24.Text
            Sheet_IT_8B_G3.Cells(6, 3).Value = frmIT_8B_G3.cmb25.Text
            Sheet_IT_8B_G3.Cells(7, 3).Value = frmIT_8B_G3.cmb26.Text
            Sheet_IT_8B_G3.Cells(8, 3).Value = frmIT_8B_G3.cmb27.Text
            Sheet_IT_8B_G3.Cells(9, 3).Value = frmIT_8B_G3.cmb28.Text
            Sheet_IT_8B_G3.Cells(10, 3).Value = frmIT_8B_G3.cmb29.Text
            Sheet_IT_8B_G3.Cells(11, 3).Value = frmIT_8B_G3.cmb30.Text

            ' Wednessday

            Sheet_IT_8B_G3.Cells(2, 4).Value = frmIT_8B_G3.cmb41.Text
            Sheet_IT_8B_G3.Cells(3, 4).Value = frmIT_8B_G3.cmb42.Text
            Sheet_IT_8B_G3.Cells(4, 4).Value = frmIT_8B_G3.cmb43.Text
            Sheet_IT_8B_G3.Cells(5, 4).Value = frmIT_8B_G3.cmb44.Text
            Sheet_IT_8B_G3.Cells(6, 4).Value = frmIT_8B_G3.cmb45.Text
            Sheet_IT_8B_G3.Cells(7, 4).Value = frmIT_8B_G3.cmb46.Text
            Sheet_IT_8B_G3.Cells(8, 4).Value = frmIT_8B_G3.cmb47.Text
            Sheet_IT_8B_G3.Cells(9, 4).Value = frmIT_8B_G3.cmb48.Text
            Sheet_IT_8B_G3.Cells(10, 4).Value = frmIT_8B_G3.cmb49.Text
            Sheet_IT_8B_G3.Cells(11, 4).Value = frmIT_8B_G3.cmb50.Text

            ' Thursday

            Sheet_IT_8B_G3.Cells(2, 5).Value = frmIT_8B_G3.cmb61.Text
            Sheet_IT_8B_G3.Cells(3, 5).Value = frmIT_8B_G3.cmb62.Text
            Sheet_IT_8B_G3.Cells(4, 5).Value = frmIT_8B_G3.cmb63.Text
            Sheet_IT_8B_G3.Cells(5, 5).Value = frmIT_8B_G3.cmb64.Text
            Sheet_IT_8B_G3.Cells(6, 5).Value = frmIT_8B_G3.cmb65.Text
            Sheet_IT_8B_G3.Cells(7, 5).Value = frmIT_8B_G3.cmb66.Text
            Sheet_IT_8B_G3.Cells(8, 5).Value = frmIT_8B_G3.cmb67.Text
            Sheet_IT_8B_G3.Cells(9, 5).Value = frmIT_8B_G3.cmb68.Text
            Sheet_IT_8B_G3.Cells(10, 5).Value = frmIT_8B_G3.cmb69.Text
            Sheet_IT_8B_G3.Cells(11, 5).Value = frmIT_8B_G3.cmb70.Text

            ' Friday

            Sheet_IT_8B_G3.Cells(2, 6).Value = frmIT_8B_G3.cmb81.Text
            Sheet_IT_8B_G3.Cells(3, 6).Value = frmIT_8B_G3.cmb82.Text
            Sheet_IT_8B_G3.Cells(4, 6).Value = frmIT_8B_G3.cmb83.Text
            Sheet_IT_8B_G3.Cells(5, 6).Value = frmIT_8B_G3.cmb84.Text
            Sheet_IT_8B_G3.Cells(6, 6).Value = frmIT_8B_G3.cmb85.Text
            Sheet_IT_8B_G3.Cells(7, 6).Value = frmIT_8B_G3.cmb86.Text
            Sheet_IT_8B_G3.Cells(8, 6).Value = frmIT_8B_G3.cmb87.Text
            Sheet_IT_8B_G3.Cells(9, 6).Value = frmIT_8B_G3.cmb88.Text
            Sheet_IT_8B_G3.Cells(10, 6).Value = frmIT_8B_G3.cmb89.Text
            Sheet_IT_8B_G3.Cells(11, 6).Value = frmIT_8B_G3.cmb90.Text


            Sheet_IT_4A_G1 = Nothing
            Sheet_IT_4A_G2 = Nothing
            Sheet_IT_4A_G3 = Nothing
            Sheet_IT_4B_G1 = Nothing
            Sheet_IT_4B_G2 = Nothing
            Sheet_IT_4B_G3 = Nothing

            Sheet_IT_6A_G1 = Nothing
            Sheet_IT_6A_G2 = Nothing
            Sheet_IT_6A_G3 = Nothing
            Sheet_IT_6B_G1 = Nothing
            Sheet_IT_6B_G2 = Nothing
            Sheet_IT_6B_G3 = Nothing

            Sheet_IT_8A_G1 = Nothing
            Sheet_IT_8A_G2 = Nothing
            Sheet_IT_8A_G3 = Nothing
            Sheet_IT_8B_G1 = Nothing
            Sheet_IT_8B_G2 = Nothing
            Sheet_IT_8B_G3 = Nothing

        End If
        ' Make sure Excel is visible and give the user control
        ' of Excel's lifetime.
        oXL.Visible = True
        oXL.UserControl = True

        ' Make sure that you release object references.
        oRng = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing
    End Sub

    Private Sub cmb103_Click(sender As Object, e As EventArgs) Handles cmb103.Click
        If cmb102.Items.Count = 0 Then
            MsgBox("Error !!! Please first select Odd or Even Semester   !!!", vbCritical)
            cmb102.Items.Clear()
        End If
    End Sub
    Private Sub cmb102_MouseClick(sender As Object, e As MouseEventArgs) Handles cmb102.MouseClick
        If cmb102.Items.Count = 0 Then
            MsgBox("Error !!!  Please select Branch and Odd or Even Semester", vbCritical)
        End If
    End Sub

    Private Sub TeachersTimeTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TeachersTimeTableToolStripMenuItem.Click
        ' Not Implemented Yet
        MsgBox("Sorry! This feature is not available yet...!!!")
    End Sub

    Private Sub ExcelExport(sender As Object, e As EventArgs) Handles ExcelToolStripMenuItem.Click
        If MsgBox("This will create seperate Excel files for each class!!! Do you still want to continue???", vbYesNo) = 6 Then
            ExcelExport()
        End If
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        ' Shows warning and resets database ie complete rollback
        If MsgBox("All data will be lost..  Be careful...!!!\n Do you still want to continue ???", vbYesNo, "Warning...!!!") = 6 Then
            initDB()
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        ' This feature is not implemented yet! 
        MsgBox("This feature has not been implemented yet!", vbInformation, "Unavailable Feature!!!")
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintForm1.PrintAction = Printing.PrintAction.PrintToPrinter
        PrintDialog1.ShowDialog()     ' Manually set page orientation to Landscape
        PrintForm1.Print()
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        ' Some problems still exist

        PrintForm1.PrintAction = Printing.PrintAction.PrintToPreview
        'PrintDialog1.ShowDialog()
        PrintForm1.Print()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If MsgBox("Your work will be lost after this!!!  Dont Forget to Save first! Do you still want to continue ???", vbYesNo Or vbCritical, "Warning...!!!") = 6 Then
            End
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub
End Class
