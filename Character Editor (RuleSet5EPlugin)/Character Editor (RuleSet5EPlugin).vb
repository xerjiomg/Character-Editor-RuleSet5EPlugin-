Imports System.Diagnostics.Eventing.Reader
Imports System.DirectoryServices.ActiveDirectory
Imports System.IO
'Imports System.Net.WebRequestMethods
Imports System.Text
Imports System.Text.Json.Nodes
Imports System.Windows
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Character_Editor__RuleSet5EPlugin_.XJ
Imports Character_Editor__RuleSet5EPlugin_.XJ.RuleSet5ECharacterEditor
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class Form1
    ' Public chara As New Character

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim s_outpoutJson As String
        Dim settingjson As New JsonSerializerSettings
        settingjson.NullValueHandling = NullValueHandling.Ignore

        s_outpoutJson = JsonConvert.SerializeObject(chara, Newtonsoft.Json.Formatting.Indented, settingjson)
        ' MsgBox(s)

        Dim path As String = "g:\temp\MyTest.Dnd5e"

        ' Create or overwrite the file.
        Dim fs As FileStream = File.Create(path)

        ' Add text to the file.
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(s_outpoutJson)
        fs.Write(info, 0, info.Length)
        fs.Close()

        'Después, para serializarlos, usa algo como esto: Dim json As String = JsonConvert.SerializeObject(laLista).
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        ' Dim ultrachara As New resistance

        'Dim fichero As String
        Dim txtFichero As String = ""
        'Dim dlAbrir As New System.Windows.Forms.OpenFileDialog

        Dim fichero As String
        Dim dlAbrir As New System.Windows.Forms.OpenFileDialog
        Dim dt_principal As New DataTable
        dlAbrir.Filter = "Archivos Dnd5e (*.Dnd5e)|*.Dnd5e|" & "Todos los archivos (*.*)|*.*"
        dlAbrir.Multiselect = False
        dlAbrir.CheckFileExists = False
        dlAbrir.Title = "Selección de fichero"
        dlAbrir.ShowDialog()
        If dlAbrir.FileName <> "" Then
            fichero = dlAbrir.FileName
            txtFichero = File.ReadAllText(fichero)
        End If
        'MsgBox(txtFichero)
        If txtFichero <> "" Then
            chara = Nothing
            chara = JsonConvert.DeserializeObject(Of Character)(txtFichero)
            cleanall()
        End If

        dg_Principal.Rows.Add("NPC", chara.NPC)
        dg_Principal.Rows.Add("Speed", chara.speed)
        dg_Principal.Rows.Add("AC", chara.ac)
        dg_Principal.Rows.Add("Hp", chara.hp)
        dg_Principal.Rows.Add("Str", chara.str)
        dg_Principal.Rows.Add("Dex", chara.dex)
        dg_Principal.Rows.Add("Con", chara.con)
        dg_Principal.Rows.Add("Int", chara.Int)
        dg_Principal.Rows.Add("Wis", chara.wis)
        dg_Principal.Rows.Add("Cha", chara.cha)
        dg_Principal.Rows.Add("Reach", chara.reach)



        For Each im_string In chara.resistance
            If im_string.ToLower = "bludgeoning" Then
                cb_resist_Bludgeoning.Checked = True
            End If
            If im_string.ToLower = "piercing" Then
                cb_resist_Piercing.Checked = True
            End If
            If im_string.ToLower = "slashing" Then
                cb_resist_Slashing.Checked = True
            End If
            If im_string.ToLower = "acid" Then
                cb_resist_Acid.Checked = True
            End If
            If im_string.ToLower = "cold" Then
                cb_resist_Cold.Checked = True
            End If
            If im_string.ToLower = "fire" Then
                cb_resist_Fire.Checked = True
            End If
            If im_string.ToLower = "force" Then
                cb_resist_Force.Checked = True
            End If
            If im_string.ToLower = "lightning" Then
                cb_resist_Lightning.Checked = True
            End If
            If im_string.ToLower = "necrotic" Then
                cb_resist_Necrotic.Checked = True
            End If
            If im_string.ToLower = "poison" Then
                cb_resist_Poison.Checked = True
            End If
            If im_string.ToLower = "psychic" Then
                cb_resist_Psychic.Checked = True
            End If
            If im_string.ToLower = "radiant" Then
                cb_resist_Radiant.Checked = True
            End If
            If im_string.ToLower = "thunder" Then
                cb_resist_Thunder.Checked = True
            End If
        Next
        For Each im_string In chara.immunity
            If im_string.ToLower = "bludgeoning" Then
                cb_immu_Bludgeoning.Checked = True
            End If
            If im_string.ToLower = "piercing" Then
                cb_immu_Piercing.Checked = True
            End If
            If im_string.ToLower = "slashing" Then
                cb_immu_Slashing.Checked = True
            End If
            If im_string.ToLower = "acid" Then
                cb_immu_Acid.Checked = True
            End If
            If im_string.ToLower = "cold" Then
                cb_immu_Cold.Checked = True
            End If
            If im_string.ToLower = "fire" Then
                cb_immu_Fire.Checked = True
            End If
            If im_string.ToLower = "force" Then
                cb_immu_Force.Checked = True
            End If
            If im_string.ToLower = "lightning" Then
                cb_immu_Lightning.Checked = True
            End If
            If im_string.ToLower = "necrotic" Then
                cb_immu_Necrotic.Checked = True
            End If
            If im_string.ToLower = "poison" Then
                cb_immu_Poison.Checked = True
            End If
            If im_string.ToLower = "psychic" Then
                cb_immu_Psychic.Checked = True
            End If
            If im_string.ToLower = "radiant" Then
                cb_immu_Radiant.Checked = True
            End If
            If im_string.ToLower = "thunder" Then
                cb_immu_Thunder.Checked = True
            End If
        Next
        For Each im_string In chara.vulnerability
            If im_string.ToLower = "bludgeoning" Then
                cb_vulne_Bludgeoning.Checked = True
            End If
            If im_string.ToLower = "piercing" Then
                cb_vulne_Piercing.Checked = True
            End If
            If im_string.ToLower = "slashing" Then
                cb_vulne_Slashing.Checked = True
            End If
            If im_string.ToLower = "acid" Then
                cb_vulne_Acid.Checked = True
            End If
            If im_string.ToLower = "cold" Then
                cb_vulne_Cold.Checked = True
            End If
            If im_string.ToLower = "fire" Then
                cb_vulne_Fire.Checked = True
            End If
            If im_string.ToLower = "force" Then
                cb_vulne_Force.Checked = True
            End If
            If im_string.ToLower = "lightning" Then
                cb_vulne_Lightning.Checked = True
            End If
            If im_string.ToLower = "necrotic" Then
                cb_vulne_Necrotic.Checked = True
            End If
            If im_string.ToLower = "poison" Then
                cb_vulne_Poison.Checked = True
            End If
            If im_string.ToLower = "psychic" Then
                cb_vulne_Psychic.Checked = True
            End If
            If im_string.ToLower = "radiant" Then
                cb_vulne_Radiant.Checked = True
            End If
            If im_string.ToLower = "thunder" Then
                cb_vulne_Thunder.Checked = True
            End If
        Next

        For Each roll In chara.saves
            lb_Saves.Items.Add(roll.name)
        Next
        For Each roll In chara.skills
            lb_Skill.Items.Add(roll.name)
        Next
        For Each roll In chara.attacks
            lb_Attacks.Items.Add(roll.name)
        Next
        For Each roll In chara.attacksDC
            lb_DCAttacks.Items.Add(roll.name)
        Next
        For Each roll In chara.healing
            lb_Heals.Items.Add(roll.name)
        Next

    End Sub


    Private Sub dg_Principal_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dg_Principal.CellValueChanged
        Select Case e.RowIndex
            Case 0
                If dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString.ToUpper <> "TRUE" And dg_Principal.Rows(0).Cells(1).Value.ToString.ToUpper <> "FALSE" Then
                    MsgBox("Invalid value (NPC only allow true or false)")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.NPC.ToString
                Else
                    chara.NPC = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString.ToLower
                End If
            Case 1
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.speed = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.speed.ToString
                End If
            Case 2
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.ac = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.ac.ToString
                End If
            Case 3
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.hp = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.hp.ToString
                End If
            Case 4
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.str = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.str.ToString
                End If
            Case 5
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.dex = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.dex.ToString
                End If
            Case 6
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.con = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.con.ToString
                End If
            Case 7
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.Int = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.Int.ToString
                End If
            Case 8
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.wis = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.wis.ToString
                End If
            Case 9
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.cha = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.cha.ToString
                End If
            Case 10
                If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                    If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                        chara.reach = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    End If
                Else
                    MsgBox("Invalid value: must be an integer")
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.reach.ToString
                End If
        End Select
    End Sub

    Private Sub lb_Skill_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_Skill.SelectedIndexChanged

        btn_Add.Text = "Add Skill Roll"
        Btn_Delete.Text = "Delete Skill Roll"
        dg_Rolls.Rows.Clear()
        lb_damages.Items.Clear()
        dg_Damages.Rows.Clear()

        If lb_Skill.SelectedIndex <> -1 Then

            Dim CBox As New DataGridViewComboBoxCell()
            dg_Rolls.Rows.Add("Type", "")
            CBox.Items.Add(chara.skills(lb_Skill.SelectedIndex).type.ToString)
            CBox.Items.Add("Public")
            CBox.Items.Add("Private")
            CBox.Items.Add("Secret")
            CBox.Items.Add("Public,GM")
            CBox.Items.Add("Private,GM")
            CBox.Items.Add("Secret,GM")
            CBox.Value = chara.skills(lb_Skill.SelectedIndex).type.ToString
            dg_Rolls.Rows(0).Cells(1) = CBox
            dg_Rolls.Rows.Add("Name", chara.skills(lb_Skill.SelectedIndex).name.ToString)
            dg_Rolls.Rows.Add("Range", "")
            dg_Rolls.Rows.Add("Roll", chara.skills(lb_Skill.SelectedIndex).roll.ToString)
            dg_Rolls.Rows.Add("Crit Multiplier", "")
            dg_Rolls.Rows.Add("Crit Range Min", "")
            dg_Rolls.Rows.Add("Info", chara.skills(lb_Skill.SelectedIndex).info.ToString)
            dg_Rolls.Rows(2).Visible = False
            dg_Rolls.Rows(4).Visible = False
            dg_Rolls.Rows(5).Visible = False

        End If
    End Sub

    Private Sub lb_Saves_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_Saves.SelectedIndexChanged
        btn_Add.Text = "Add Save Roll"
        Btn_Delete.Text = "Delete Save Roll"
        dg_Rolls.Rows.Clear()
        lb_damages.Items.Clear()
        dg_Damages.Rows.Clear()
        If lb_Saves.SelectedIndex <> -1 Then
            Dim CBox As New DataGridViewComboBoxCell()
            dg_Rolls.Rows.Add("Type", "")
            CBox.Items.Add(chara.saves(lb_Saves.SelectedIndex).type.ToString)
            CBox.Items.Add("Public")
            CBox.Items.Add("Private")
            CBox.Items.Add("Secret")
            CBox.Items.Add("Public,GM")
            CBox.Items.Add("Private,GM")
            CBox.Items.Add("Secret,GM")
            CBox.Value = chara.saves(lb_Saves.SelectedIndex).type.ToString
            dg_Rolls.Rows(0).Cells(1) = CBox
            dg_Rolls.Rows.Add("Name", chara.saves(lb_Saves.SelectedIndex).name.ToString)
            dg_Rolls.Rows.Add("Range", "")
            dg_Rolls.Rows.Add("Roll", chara.saves(lb_Saves.SelectedIndex).roll.ToString)
            dg_Rolls.Rows.Add("Crit Multiplier", "")
            dg_Rolls.Rows.Add("Crit Range Min", "")
            dg_Rolls.Rows.Add("Info", chara.saves(lb_Saves.SelectedIndex).info.ToString)
            dg_Rolls.Rows(2).Visible = False
            dg_Rolls.Rows(4).Visible = False
            dg_Rolls.Rows(5).Visible = False
        End If

    End Sub

    Private Sub lb_Attacks_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_Attacks.SelectedIndexChanged
        btn_Add.Text = "Add Attack"
        Btn_Delete.Text = "Delete Attack"
        dg_Rolls.Rows.Clear()
        lb_damages.Items.Clear()
        dg_Damages.Rows.Clear()

        If lb_Attacks.SelectedIndex <> -1 Then
            Dim CBox As New DataGridViewComboBoxCell()
            dg_Rolls.Rows.Add("Type", "")
            CBox.Items.Add(chara.attacks(lb_Attacks.SelectedIndex).type.ToString)
            CBox.Items.Add("Melee")
            CBox.Items.Add("Ranged")
            CBox.Items.Add("Magic")

            CBox.Value = chara.attacks(lb_Attacks.SelectedIndex).type.ToString
            dg_Rolls.Rows(0).Cells(1) = CBox
            dg_Rolls.Rows.Add("Name", chara.attacks(lb_Attacks.SelectedIndex).name.ToString)
            dg_Rolls.Rows.Add("Range", chara.attacks(lb_Attacks.SelectedIndex).range.ToString)
            dg_Rolls.Rows.Add("Roll", chara.attacks(lb_Attacks.SelectedIndex).roll.ToString)
            dg_Rolls.Rows.Add("Crit Multiplier", chara.attacks(lb_Attacks.SelectedIndex).critmultip.ToString)
            dg_Rolls.Rows.Add("Crit Range Min", chara.attacks(lb_Attacks.SelectedIndex).critrangemin.ToString)
            dg_Rolls.Rows.Add("Info", chara.attacks(lb_Attacks.SelectedIndex).info.ToString)
            Dim tlink As Roll
            If chara.attacks(lb_Attacks.SelectedIndex).link IsNot Nothing Then
                tlink = chara.attacks(lb_Attacks.SelectedIndex).link
                lb_damages.Items.Add(tlink.name)
                Dim i As Integer = 2
                While tlink.link IsNot Nothing
                    Dim s_tlink As String = ""
                    If tlink.name = tlink.link.name Then
                        s_tlink = " (" + i.ToString + ")"
                    End If
                    lb_damages.Items.Add(tlink.link.name + s_tlink)
                    tlink = tlink.link
                    i = i + 1
                End While
            End If
        End If

    End Sub

    Private Sub lb_DCAttacks_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_DCAttacks.SelectedIndexChanged
        btn_Add.Text = "Add DC Attack"
        Btn_Delete.Text = "Delete DC Attack"
        dg_Rolls.Rows.Clear()
        lb_damages.Items.Clear()
        dg_Damages.Rows.Clear()

        If lb_DCAttacks.SelectedIndex <> -1 Then
            Dim CBox As New DataGridViewComboBoxCell()
            dg_Rolls.Rows.Add("Type", "")
            CBox.Items.Add(chara.attacksDC(lb_DCAttacks.SelectedIndex).type.ToString)
            CBox.Items.Add("Melee")
            CBox.Items.Add("Ranged")
            CBox.Items.Add("Magic")

            CBox.Value = chara.attacksDC(lb_DCAttacks.SelectedIndex).type.ToString
            dg_Rolls.Rows(0).Cells(1) = CBox
            dg_Rolls.Rows.Add("Name", chara.attacksDC(lb_DCAttacks.SelectedIndex).name.ToString)
            dg_Rolls.Rows.Add("Range", chara.attacksDC(lb_DCAttacks.SelectedIndex).range.ToString)
            dg_Rolls.Rows.Add("Roll", chara.attacksDC(lb_DCAttacks.SelectedIndex).roll.ToString)
            dg_Rolls.Rows.Add("Crit Multiplier", chara.attacksDC(lb_DCAttacks.SelectedIndex).critmultip.ToString)
            dg_Rolls.Rows.Add("Crit Range Min", chara.attacksDC(lb_DCAttacks.SelectedIndex).critrangemin.ToString)
            dg_Rolls.Rows.Add("Info", chara.attacksDC(lb_DCAttacks.SelectedIndex).info.ToString)
            dg_Damages.Rows.Add()
            Dim tlink As Roll
            If chara.attacksDC(lb_DCAttacks.SelectedIndex).link IsNot Nothing Then
                tlink = chara.attacksDC(lb_DCAttacks.SelectedIndex).link
                lb_damages.Items.Add(tlink.name)
                Dim i As Integer = 2
                While tlink.link IsNot Nothing
                    Dim s_tlink As String = ""
                    If tlink.name = tlink.link.name Then
                        s_tlink = " (" + i.ToString + ")"
                    End If
                    lb_damages.Items.Add(tlink.link.name + s_tlink)
                    tlink = tlink.link
                    i = i + 1
                End While
            End If
        End If


    End Sub

    Private Sub lb_Heals_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_Heals.SelectedIndexChanged

        btn_Add.Text = "Add Heal"
        Btn_Delete.Text = "Delete Heal"
        dg_Rolls.Rows.Clear()
        lb_damages.Items.Clear()
        dg_Damages.Rows.Clear()
        If lb_Heals.SelectedIndex <> -1 Then
            Dim CBox As New DataGridViewComboBoxCell()
            dg_Rolls.Rows.Add("Type", "")
            CBox.Items.Add(chara.healing(lb_Heals.SelectedIndex).type.ToString)
            CBox.Items.Add("Magic")
            dg_Rolls.Rows(0).Cells(1) = CBox
            CBox.Value = chara.healing(lb_Heals.SelectedIndex).type.ToString
            dg_Rolls.Rows.Add("Name", chara.healing(lb_Heals.SelectedIndex).name.ToString)
            dg_Rolls.Rows.Add("Range", chara.healing(lb_Heals.SelectedIndex).range.ToString)
            dg_Rolls.Rows.Add("Roll", chara.healing(lb_Heals.SelectedIndex).roll.ToString)
            dg_Rolls.Rows.Add("Crit Multiplier", "")
            dg_Rolls.Rows.Add("Crit Range Min", "")
            dg_Rolls.Rows.Add("Info", chara.healing(lb_Heals.SelectedIndex).info.ToString)
            dg_Rolls.Rows(4).Visible = False
            dg_Rolls.Rows(5).Visible = False
        End If

    End Sub

    Private Sub lb_Skill_GotFocus(sender As Object, e As EventArgs) Handles lb_Skill.GotFocus
        For Each lbc As ListBox In panel_lists.Controls.OfType(Of ListBox)
            If lbc.Tag <> sender.Tag Then
                lbc.ClearSelected()
            End If
        Next
    End Sub
    Private Sub lb_Attacks_GotFocus(sender As Object, e As EventArgs) Handles lb_Attacks.GotFocus
        For Each lbc As ListBox In panel_lists.Controls.OfType(Of ListBox)
            If lbc.Tag <> sender.Tag Then
                lbc.ClearSelected()
            End If
        Next
    End Sub
    Private Sub lb_DCAttacks_GotFocus(sender As Object, e As EventArgs) Handles lb_DCAttacks.GotFocus
        For Each lbc As ListBox In panel_lists.Controls.OfType(Of ListBox)
            If lbc.Tag <> sender.Tag Then
                lbc.ClearSelected()
            End If
        Next
    End Sub
    Private Sub lb_Heals_GotFocus(sender As Object, e As EventArgs) Handles lb_Heals.GotFocus
        For Each lbc As ListBox In panel_lists.Controls.OfType(Of ListBox)
            If lbc.Tag <> sender.Tag Then
                lbc.ClearSelected()
            End If
        Next
    End Sub
    Private Sub lb_Saves_GotFocus(sender As Object, e As EventArgs) Handles lb_Saves.GotFocus
        For Each lbc As ListBox In panel_lists.Controls.OfType(Of ListBox)
            If lbc.Tag <> sender.Tag Then
                lbc.ClearSelected()
            End If
        Next
    End Sub

    Private Sub lb_damages_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_damages.SelectedIndexChanged
        dg_Damages.Rows.Clear()
        Dim dlink As Roll
        If lb_DCAttacks.SelectedIndex <> -1 Then
            dlink = chara.attacksDC(lb_DCAttacks.SelectedIndex).link
            For i As Integer = 0 To lb_damages.SelectedIndex - 1
                dlink = dlink.link
            Next
            Dim CBox As New DataGridViewComboBoxCell()
            dg_Damages.Rows.Add("Type", "")
            CBox.Items.Add(dlink.type.ToString)
            CBox.Items.Add("")
            For Each s_damage As String In damagesArray
                CBox.Items.Add(s_damage)
            Next
            dg_Damages.Rows(0).Cells(1) = CBox
            CBox.Value = dlink.type.ToString
            dg_Damages.Rows.Add("Name", dlink.name.ToString)
            dg_Damages.Rows.Add("Range", dlink.range.ToString)
            dg_Damages.Rows.Add("Roll", dlink.roll.ToString)
            dg_Rolls.Rows.Add("Crit Multiplier", dlink.critmultip.ToString)
            dg_Rolls.Rows.Add("Crit Range Min", dlink.critrangemin.ToString)
            dg_Damages.Rows.Add("Info", dlink.info.ToString)
            dg_Rolls.Rows(4).Visible = False
            dg_Rolls.Rows(5).Visible = False

        ElseIf lb_Attacks.SelectedIndex <> -1 Then
            dlink = chara.attacks(lb_Attacks.SelectedIndex).link
            For i As Integer = 0 To lb_damages.SelectedIndex - 1
                dlink = dlink.link
            Next
            Dim CBox As New DataGridViewComboBoxCell()
            dg_Damages.Rows.Add("Type", "")
            CBox.Items.Add("")
            For Each s_damage As String In damagesArray
                CBox.Items.Add(s_damage)
            Next
            dg_Damages.Rows(0).Cells(1) = CBox
            CBox.Value = dlink.type.ToString
            dg_Damages.Rows.Add("Name", dlink.name.ToString)
            dg_Damages.Rows.Add("Range", dlink.range.ToString)
            dg_Damages.Rows.Add("Roll", dlink.roll.ToString)
            dg_Rolls.Rows.Add("Crit Multiplier", dlink.critmultip.ToString)
            dg_Rolls.Rows.Add("Crit Range Min", dlink.critrangemin.ToString)
            dg_Damages.Rows.Add("Info", dlink.info.ToString)
            dg_Rolls.Rows(4).Visible = False
            dg_Rolls.Rows(5).Visible = False
        Else
            dlink = Nothing
            MsgBox("Attack list item not selected")
        End If


    End Sub

    Private Sub dg_Rolls_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dg_Rolls.CellValueChanged

        Dim s_caller As String = ""
        Dim l_caller As ListBox = Nothing
        Dim b_doChanges As Boolean = False
        Dim i_caller As Integer = -1

        '' ¿Qué listbox me está llamando?
        For Each l_caller In panel_lists.Controls.OfType(Of ListBox)
            If l_caller.SelectedIndex <> -1 Then
                s_caller = l_caller.Name
                i_caller = l_caller.SelectedIndex
            End If
        Next

        If s_caller <> "" Then


            Select Case e.RowIndex
                Case 0
                    chara.attacks(i_caller).type = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                Case 1
                    If dg_Rolls.Rows(e.RowIndex).Cells(1).Value <> 0 And dg_Rolls.Rows(e.RowIndex).Cells(1).Value <> Nothing Then
                        chara.attacks(i_caller).name = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                    Else
                        dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                        MsgBox("The name field cannot be empty")
                    End If
                Case 2
                    If dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Contains("/") Then
                        Dim s_ranges As String() = dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Split("/")
                        If s_ranges.Length = 2 And IsNumeric(s_ranges(0)) And IsNumeric(s_ranges(1)) And (s_ranges(0) Mod 1) = 0 And (s_ranges(1) Mod 1) = 0 Then
                            chara.attacks(i_caller).range = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                        End If
                    Else
                        If dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString = "" Or dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString = Nothing Then
                            chara.attacks(i_caller).range = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                        Else
                            dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                            MsgBox("Must have two integer numbers splited by: / ")
                        End If
                    End If
                Case 3
                    If s_caller = "lb_DCAttacks" Then
                        If dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString = "100" Then
                            chara.attacks(i_caller).roll = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                        Else
                            Dim s_roll_Dc As String() = {""}
                            If dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Contains("/") Then
                                s_roll_Dc = dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Split("/")
                            End If
                            If s_roll_Dc.Length = 3 Then
                                If IsNumeric(s_roll_Dc(0)) And (s_roll_Dc(0) Mod 1) = 0 And s_validStats.Contains(s_roll_Dc(1).ToUpper) And s_validDc3rdvar.Contains(s_roll_Dc(2).ToUpper) Then
                                    chara.attacks(i_caller).roll = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                                Else
                                    dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                    MsgBox("The value must be 100 or 3 values split by (/): Number/" + s_validStats + "/" + s_validDc3rdvar)
                                End If
                            Else
                                dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                MsgBox("The value must be 100 or 3 values split by (/): Number/" + s_validStats + "/" + s_validDc3rdvar)
                            End If
                        End If
                    Else
                        Dim s_validrol As String = isValidrole(dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString)
                        If s_validrol = "" Then
                            chara.attacks(i_caller).roll = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                        Else
                            dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                            MsgBox(s_validrol)
                        End If
                    End If
                Case 4
                    If IsNumeric(dg_Rolls.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Rolls.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.attacks(i_caller).critmultip = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                        End If
                    Else
                        dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                        MsgBox("Invalid value: must be an integer")

                    End If
                Case 5
                    If IsNumeric(dg_Rolls.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Rolls.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            If dg_Rolls.Rows(e.RowIndex).Cells(1).Value > 0 & dg_Rolls.Rows(e.RowIndex).Cells(1).Value < 21 Then
                                chara.attacks(i_caller).critrangemin = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                            End If
                        End If
                    Else
                        dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                        MsgBox("It must be a value between 1 and 20")
                    End If
                Case 6

            End Select
        Else
            MsgBox("There is no selected list")
        End If

    End Sub

    Private Sub dg_Rolls_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dg_Rolls.CellValidating
        dg_Rolls.Tag = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
    End Sub


    Private Sub dg_Rolls_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dg_Rolls.DataError
        MsgBox(e.Exception.ToString)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim s_outpoutJson As String
        Dim settingjson As New JsonSerializerSettings
        settingjson.NullValueHandling = NullValueHandling.Ignore
        s_outpoutJson = JsonConvert.SerializeObject(chara, Newtonsoft.Json.Formatting.Indented, settingjson)
        From_ShowJson.Show()
        From_ShowJson.s_JsonText.Text = s_outpoutJson
    End Sub
End Class

