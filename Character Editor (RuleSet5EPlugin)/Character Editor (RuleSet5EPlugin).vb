
Imports System.Buffers.Text
Imports System.Drawing.Text
Imports System.IO
Imports System.Net.Http
Imports System.Net.Security

'Imports System.Net.WebRequestMethods
Imports System.Text
Imports System.Text.Unicode
Imports Character_Editor__RuleSet5EPlugin_.XJ.RuleSet5ECharacterEditor
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

'' VERSIÓN 1.2.0

Public Class Form1
    ' Public chara As New Character

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim s_outpoutJson As String
        Dim settingjson As New JsonSerializerSettings

        settingjson.NullValueHandling = NullValueHandling.Ignore
        settingjson.DefaultValueHandling = DefaultValueHandling.Ignore

        s_outpoutJson = JsonConvert.SerializeObject(chara, Newtonsoft.Json.Formatting.Indented, settingjson)


        ' Dim txtFichero As String = ""
        'Dim dlAbrir As New System.Windows.Forms.OpenFileDialog

        Dim dlAbrir As New System.Windows.Forms.SaveFileDialog
        Dim dt_principal As New DataTable
        dlAbrir.InitialDirectory = T_DefaultCPath.Text
        dlAbrir.Filter = "Archivos Dnd5e (*.Dnd5e)|*.Dnd5e|" & "Todos los archivos (*.*)|*.*"
        'dlAbrir.Multiselect = False
        dlAbrir.CheckFileExists = False
        dlAbrir.Title = "Selección de fichero"
        dlAbrir.ShowDialog()
        If dlAbrir.FileName <> "" Then
            Dim path As String = dlAbrir.FileName
            Dim fs As FileStream = File.Create(path)
            ' Add text to the file.
            Dim info As Byte() = New UTF8Encoding(True).GetBytes(s_outpoutJson)
            fs.Write(info, 0, info.Length)
            fs.Close()

            Dim _tsCreature() As String = dlAbrir.FileName.ToString.Substring(InStrRev(dlAbrir.FileName.ToString, "\"), dlAbrir.FileName.ToString.Length - InStrRev(dlAbrir.FileName.ToString, "\") - 1).Split(".")
            l_creaturename.Text = _tsCreature(0)

        Else
        End If


        ' Create or overwrite the file.



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try


            ' Dim ultrachara As New resistance

            'Dim fichero As String
            Dim txtFichero As String = ""
            'Dim dlAbrir As New System.Windows.Forms.OpenFileDialog

            Dim fichero As String
            Dim dlAbrir As New System.Windows.Forms.OpenFileDialog
            Dim dt_principal As New DataTable
            dlAbrir.InitialDirectory = T_DefaultCPath.Text
            dlAbrir.Filter = "File Dnd5e (*.Dnd5e)|*.Dnd5e|" & "Todos los archivos (*.*)|*.*"
            dlAbrir.Multiselect = False
            dlAbrir.CheckFileExists = False
            dlAbrir.Title = "Select File"
            dlAbrir.ShowDialog()
            If dlAbrir.FileName <> "" And System.IO.File.Exists(dlAbrir.FileName) Then
                fichero = dlAbrir.FileName
                txtFichero = File.ReadAllText(fichero)
                Dim _tsCreature() As String = dlAbrir.SafeFileName.ToString.Split(".")
                l_creaturename.Text = _tsCreature(0)

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
                dg_Principal.Rows.Add("lv", chara.lv)
                dg_Principal.Rows.Add("var1", chara.var1)
                dg_Principal.Rows.Add("var2", chara.var2)
                dg_Principal.Rows.Add("var3", chara.var3)
                '' Cambio chorizero de color chorra
                If dg_Principal.Rows.Count <> 0 Then
                    changecolors()
                End If

                b_checkedchanges = False

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
                    If im_string.ToLower = "magic bludgeoning" Then
                        cb_resist_Magic0Bludgeoning.Checked = True
                    End If
                    If im_string.ToLower = "magic piercing" Then
                        cb_resist_Magic0Piercing.Checked = True
                    End If
                    If im_string.ToLower = "magic slashing" Then
                        cb_resist_Magic0Slashing.Checked = True
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
                    If im_string.ToLower = "magic bludgeoning" Then
                        cb_immu_Magic0Bludgeoning.Checked = True
                    End If
                    If im_string.ToLower = "magic piercing" Then
                        cb_immu_Magic0Piercing.Checked = True
                    End If
                    If im_string.ToLower = "magic slashing" Then
                        cb_immu_Magic0Slashing.Checked = True
                    End If
                    If im_string.ToLower = "critical" Then
                        cb_immu_Critical.Checked = True
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
                    If im_string.ToLower = "magic bludgeoning" Then
                        cb_vulne_Magic0Bludgeoning.Checked = True
                    End If
                    If im_string.ToLower = "magic piercing" Then
                        cb_vulne_Magic0Piercing.Checked = True
                    End If
                    If im_string.ToLower = "magic slashing" Then
                        cb_vulne_Magic0Slashing.Checked = True
                    End If
                Next

                b_checkedchanges = True

                LoadSavesList()

                LoadSkillsList()

                For Each roll In chara.attacks
                    lb_Attacks.Items.Add(roll.name)
                Next
                For Each roll In chara.attacksDC
                    lb_DCAttacks.Items.Add(roll.name)
                Next
                For Each roll In chara.healing
                    lb_Heals.Items.Add(roll.name)
                Next
            Else
                MsgBox("The file does not exist", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
            End If
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub


    Private Sub dg_Principal_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dg_Principal.CellValueChanged
        Try
            If e.RowIndex <> -1 Then
                If dg_Principal.Rows(e.RowIndex).Cells(1).Value = Nothing Then
                    dg_Principal.Rows(e.RowIndex).Cells(1).Value = String.Empty
                End If
            End If

            Select Case e.RowIndex
                Case 0
                    If dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString.ToUpper <> "TRUE" And dg_Principal.Rows(0).Cells(1).Value.ToString.ToUpper <> "FALSE" Then
                        MsgBox("Invalid value (NPC only allow true or false)", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
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
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.speed.ToString
                    End If
                Case 2
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.ac = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.ac.ToString
                    End If
                Case 3
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.hp = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.hp.ToString
                    End If
                Case 4
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.str = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.str.ToString
                    End If
                Case 5
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.dex = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.dex.ToString
                    End If
                Case 6
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.con = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.con.ToString
                    End If
                Case 7
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.Int = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.Int.ToString
                    End If
                Case 8
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.wis = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.wis.ToString
                    End If
                Case 9
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.cha = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.cha.ToString
                    End If
                Case 10
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.reach = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.reach.ToString
                    End If
                Case 11
                    If IsNumeric(dg_Principal.Rows(e.RowIndex).Cells(1).Value) Then
                        If (dg_Principal.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                            chara.lv = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                        End If
                    Else
                        MsgBox("Invalid value: must be an integer", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.lv.ToString
                    End If
                Case 12
                    Dim tempIsvalidrole As String = isValidrole(dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString.Replace(" ", ""))
                    If tempIsvalidrole = "" Or tempIsvalidrole = "The roll field cannot be empty" Then
                        chara.var1 = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    Else
                        MsgBox(tempIsvalidrole, MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.var1.ToString
                    End If
                Case 13
                    Dim tempIsvalidrole As String = isValidrole(dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString.Replace(" ", ""))
                    If tempIsvalidrole = "" Or tempIsvalidrole = "The roll field cannot be empty" Then
                        chara.var2 = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    Else
                        MsgBox(tempIsvalidrole, MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.var2.ToString
                    End If

                Case 14
                    Dim tempIsvalidrole As String = isValidrole(dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString.Replace(" ", ""))
                    If tempIsvalidrole = "" Or tempIsvalidrole = "The roll field cannot be empty" Then
                        chara.var3 = dg_Principal.Rows(e.RowIndex).Cells(1).Value.ToString
                    Else
                        MsgBox(tempIsvalidrole, MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                        dg_Principal.Rows(e.RowIndex).Cells(1).Value = chara.var3.ToString
                    End If

            End Select

            '' Cambio chorizero de color chorra
            If dg_Principal.Rows.Count <> 0 Then
                changecolors()
            End If

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub lb_Skill_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_Skill.SelectedIndexChanged
        Try


            b_rolldgcanedit = False

            dg_Rolls.Rows.Clear()
            lb_damages.Items.Clear()
            dg_Damages.Rows.Clear()

            If lb_Skill.SelectedItems.Count > 0 Then

                Dim CBox As New DataGridViewComboBoxCell()
                ''Chorradas del color
                CBox.Style.BackColor = BackColor
                CBox.FlatStyle = FlatStyle.Flat

                dg_Rolls.Rows.Add("Type", "")
                CBox.Items.Add(chara.skills(lb_Skill.SelectedItems(0).Index).type.ToString)
                CBox.Items.Add("Public")
                CBox.Items.Add("Private")
                CBox.Items.Add("Secret")
                CBox.Items.Add("Public,GM")
                CBox.Items.Add("Private,GM")
                CBox.Items.Add("Secret,GM")
                CBox.Value = chara.skills(lb_Skill.SelectedItems(0).Index).type.ToString
                dg_Rolls.Rows(0).Cells(1) = CBox
                dg_Rolls.Rows.Add("Name", chara.skills(lb_Skill.SelectedItems(0).Index).name.ToString)
                dg_Rolls.Rows.Add("Range", "")
                dg_Rolls.Rows.Add("Roll", chara.skills(lb_Skill.SelectedItems(0).Index).roll.ToString)
                dg_Rolls.Rows.Add("Crit Multiplier", "")
                dg_Rolls.Rows.Add("Crit Range Min", "")
                dg_Rolls.Rows.Add("Info", chara.skills(lb_Skill.SelectedItems(0).Index).info.ToString)
                dg_Rolls.Rows.Add("futureUse_icon", chara.skills(lb_Skill.SelectedItems(0).Index).futureUse_icon.ToString)
                dg_Rolls.Rows.Add("menuUI", chara.skills(lb_Skill.SelectedItems(0).Index).menuUI.ToString)
                dg_Rolls.Rows(2).Visible = False
                dg_Rolls.Rows(4).Visible = False
                dg_Rolls.Rows(5).Visible = False
                dg_Rolls.Rows(7).Visible = False

                checkProficencyEvent = False

                If chara.skills(lb_Skill.SelectedItems(0).Index).roll.Contains("{pb}") Then
                    c_profSkill.Checked = True
                Else
                    c_profSkill.Checked = False
                End If
                If chara.skills(lb_Skill.SelectedItems(0).Index).roll.Contains("{ex}") Then
                    c_experSkill.Checked = True
                Else
                    c_experSkill.Checked = False
                End If
                If chara.skills(lb_Skill.SelectedItems(0).Index).roll.Contains("{ph}") Then
                    c_profhalfSkill.Checked = True
                Else
                    c_profhalfSkill.Checked = False
                End If
                checkProficencyEvent = True
            End If
                b_rolldgcanedit = True
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub lb_Saves_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_Saves.SelectedIndexChanged
        Try


            b_rolldgcanedit = False
            dg_Rolls.Rows.Clear()
            lb_damages.Items.Clear()
            dg_Damages.Rows.Clear()
            If lb_Saves.SelectedItems.Count > 0 Then
                Dim CBox As New DataGridViewComboBoxCell()
                ''Chorradas del color
                CBox.Style.BackColor = BackColor
                CBox.FlatStyle = FlatStyle.Flat

                dg_Rolls.Rows.Add("Type", "")
                CBox.Items.Add(chara.saves(lb_Saves.SelectedItems(0).Index).type.ToString)
                CBox.Items.Add("Public")
                CBox.Items.Add("Private")
                CBox.Items.Add("Secret")
                CBox.Items.Add("Public,GM")
                CBox.Items.Add("Private,GM")
                CBox.Items.Add("Secret,GM")
                CBox.Value = chara.saves(lb_Saves.SelectedItems(0).Index).type.ToString
                dg_Rolls.Rows(0).Cells(1) = CBox
                dg_Rolls.Rows.Add("Name", chara.saves(lb_Saves.SelectedItems(0).Index).name.ToString)
                dg_Rolls.Rows.Add("Range", "")
                dg_Rolls.Rows.Add("Roll", chara.saves(lb_Saves.SelectedItems(0).Index).roll.ToString)
                dg_Rolls.Rows.Add("Crit Multiplier", "")
                dg_Rolls.Rows.Add("Crit Range Min", "")
                dg_Rolls.Rows.Add("Info", chara.saves(lb_Saves.SelectedItems(0).Index).info.ToString)
                dg_Rolls.Rows.Add("futureUse_icon", chara.saves(lb_Saves.SelectedItems(0).Index).futureUse_icon.ToString)
                dg_Rolls.Rows.Add("menuUI", chara.saves(lb_Saves.SelectedItems(0).Index).menuUI.ToString)
                dg_Rolls.Rows(2).Visible = False
                dg_Rolls.Rows(4).Visible = False
                dg_Rolls.Rows(5).Visible = False
                dg_Rolls.Rows(7).Visible = False

                checkProficencyEvent = False
                If chara.saves(lb_Saves.SelectedItems(0).Index).roll.Contains("{pb}") Then
                    c_profSaves.Checked = True
                Else
                    c_profSaves.Checked = False
                End If
                checkProficencyEvent = True

            End If
            b_rolldgcanedit = True
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub lb_Attacks_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_Attacks.SelectedIndexChanged
        Try

            b_rolldgcanedit = False

            dg_Rolls.Rows.Clear()
            lb_damages.Items.Clear()
            dg_Damages.Rows.Clear()

            If lb_Attacks.SelectedIndex <> -1 Then
                Dim CBox As New DataGridViewComboBoxCell()
                ''Chorradas del color
                CBox.Style.BackColor = BackColor
                CBox.FlatStyle = FlatStyle.Flat

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
                dg_Rolls.Rows.Add("futureUse_icon", chara.attacks(lb_Attacks.SelectedIndex).futureUse_icon.ToString)
                dg_Rolls.Rows.Add("menuUI", chara.attacks(lb_Attacks.SelectedIndex).menuUI.ToString)
                If dg_Rolls.Rows(2).Cells(1).Value = "" Then
                    dg_Rolls.Rows(2).Cells(1).Value = "5/5"
                End If
                If dg_Rolls.Rows(4).Cells(1).Value = "" Then
                    dg_Rolls.Rows(4).Cells(1).Value = "2"
                End If
                If dg_Rolls.Rows(5).Cells(1).Value = "" Then
                    dg_Rolls.Rows(5).Cells(1).Value = "20"
                End If

                '''daños uraños
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
            b_rolldgcanedit = True
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub lb_DCAttacks_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_DCAttacks.SelectedIndexChanged
        Try

            b_rolldgcanedit = False

            dg_Rolls.Rows.Clear()
            lb_damages.Items.Clear()
            dg_Damages.Rows.Clear()

            If lb_DCAttacks.SelectedIndex <> -1 Then
                Dim CBox As New DataGridViewComboBoxCell()
                ''Chorradas del color
                CBox.Style.BackColor = BackColor
                CBox.FlatStyle = FlatStyle.Flat

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
                dg_Rolls.Rows.Add("futureUse_icon", chara.attacksDC(lb_DCAttacks.SelectedIndex).futureUse_icon.ToString)
                dg_Rolls.Rows.Add("menuUI", chara.attacksDC(lb_DCAttacks.SelectedIndex).menuUI.ToString)
                dg_Rolls.Rows(4).Visible = False
                dg_Rolls.Rows(5).Visible = False
                If dg_Rolls.Rows(2).Cells(1).Value = "" Then
                    dg_Rolls.Rows(2).Cells(1).Value = "5/5"
                End If


                '' Añadir daños huraños
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
            b_rolldgcanedit = True

        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub lb_Heals_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_Heals.SelectedIndexChanged
        Try


            b_rolldgcanedit = False

            dg_Rolls.Rows.Clear()
            lb_damages.Items.Clear()
            dg_Damages.Rows.Clear()
            If lb_Heals.SelectedIndex <> -1 Then
                Dim CBox As New DataGridViewComboBoxCell()
                ''Chorradas del color
                CBox.Style.BackColor = BackColor
                CBox.FlatStyle = FlatStyle.Flat

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
                dg_Rolls.Rows.Add("menuUI", chara.healing(lb_Heals.SelectedIndex).menuUI.ToString)
                dg_Rolls.Rows(4).Visible = False
                dg_Rolls.Rows(5).Visible = False
                If dg_Rolls.Rows(2).Cells(1).Value = "" Then
                    dg_Rolls.Rows(2).Cells(1).Value = "5/5"
                End If
            End If
            b_rolldgcanedit = True
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub lb_Skill_GotFocus(sender As Object, e As EventArgs) Handles lb_Skill.GotFocus
        Try
            For Each lbc As ListBox In alllists
                If lbc.Tag <> sender.Tag Then
                    lbc.ClearSelected()
                End If
            Next

            lb_Saves.SelectedItems.Clear()

        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub
    Private Sub lb_Attacks_GotFocus(sender As Object, e As EventArgs) Handles lb_Attacks.GotFocus
        Try
            For Each lbc As ListBox In alllists
                If lbc.Tag <> sender.Tag Then
                    lbc.ClearSelected()
                End If
            Next
            lb_Skill.SelectedItems.Clear()
            lb_Saves.SelectedItems.Clear()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub lb_DCAttacks_GotFocus(sender As Object, e As EventArgs) Handles lb_DCAttacks.GotFocus
        Try
            For Each lbc As ListBox In alllists
                If lbc.Tag <> sender.Tag Then
                    lbc.ClearSelected()
                End If
            Next
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub
    Private Sub lb_Heals_GotFocus(sender As Object, e As EventArgs) Handles lb_Heals.GotFocus
        Try
            For Each lbc As ListBox In alllists
                If lbc.Tag <> sender.Tag Then
                    lbc.ClearSelected()
                End If
            Next


        Catch ex As Exception

        End Try
    End Sub
    Private Sub lb_Saves_GotFocus(sender As Object, e As EventArgs) Handles lb_Saves.GotFocus
        Try
            For Each lbc As ListBox In alllists
                If lbc.Tag <> sender.Tag Then
                    lbc.ClearSelected()
                End If
            Next

            lb_Skill.SelectedItems.Clear()

        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub lb_damages_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_damages.SelectedIndexChanged
        Try
            b_damagedgcanedit = False
            dg_Damages.Rows.Clear()
            Dim dlink As Roll
            If lb_DCAttacks.SelectedIndex <> -1 Then
                dlink = chara.attacksDC(lb_DCAttacks.SelectedIndex).link
                For i As Integer = 0 To lb_damages.SelectedIndex - 1
                    dlink = dlink.link
                Next
                Dim CBox As New DataGridViewComboBoxCell()
                ''Chorradas del color
                CBox.Style.BackColor = BackColor
                CBox.FlatStyle = FlatStyle.Flat

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
                dg_Damages.Rows.Add("Crit Multiplier", dlink.critmultip.ToString)
                dg_Damages.Rows.Add("Crit Range Min", dlink.critrangemin.ToString)
                dg_Damages.Rows.Add("Info", dlink.info.ToString)
                dg_Damages.Rows(2).Visible = False
                dg_Damages.Rows(4).Visible = False
                dg_Damages.Rows(5).Visible = False

            ElseIf lb_Attacks.SelectedIndex <> -1 Then
                dlink = chara.attacks(lb_Attacks.SelectedIndex).link
                For i As Integer = 0 To lb_damages.SelectedIndex - 1
                    dlink = dlink.link
                Next
                Dim CBox As New DataGridViewComboBoxCell()
                ''Chorradas del color
                CBox.Style.BackColor = BackColor
                CBox.FlatStyle = FlatStyle.Flat

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
                dg_Damages.Rows.Add("Crit Multiplier", dlink.critmultip.ToString)
                dg_Damages.Rows.Add("Crit Range Min", dlink.critrangemin.ToString)
                dg_Damages.Rows.Add("Info", dlink.info.ToString)
                dg_Damages.Rows(2).Visible = False
                dg_Damages.Rows(4).Visible = False
                dg_Damages.Rows(5).Visible = False

            Else
                dlink = Nothing
                MsgBox("Attack list item not selected", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
            End If

            b_damagedgcanedit = True
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub dg_Rolls_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dg_Rolls.CellValueChanged

        If e.RowIndex <> -1 Then
            If dg_Rolls.Rows(e.RowIndex).Cells(1).Value = Nothing Then
                dg_Rolls.Rows(e.RowIndex).Cells(1).Value = String.Empty
            End If
        End If

        Try

            If b_rolldgcanedit = True Then

                Dim s_caller As String = ""
                Dim l_caller As ListBox = Nothing
                Dim b_doChanges As Boolean = False
                Dim i_caller As Integer = -1

                '' ¿Qué listbox me está llamando?
                For Each l_caller In alllists
                    If l_caller.SelectedIndex <> -1 Then
                        s_caller = l_caller.Name
                        i_caller = l_caller.SelectedIndex
                    End If
                Next

                If lb_Skill.SelectedItems.Count > 0 Then
                    s_caller = lb_Skill.Name
                    i_caller = lb_Skill.SelectedItems(0).Index
                End If
                If lb_Saves.SelectedItems.Count > 0 Then
                    s_caller = lb_Saves.Name
                    i_caller = lb_Saves.SelectedItems(0).Index
                End If


                Dim tempRoll As New Roll
                Select Case s_caller
                    Case "lb_Attacks"
                        tempRoll = chara.attacks(i_caller)
                    Case "lb_DCAttacks"
                        tempRoll = chara.attacksDC(i_caller)
                    Case "lb_Heals"
                        tempRoll = chara.healing(i_caller)
                    Case "lb_Skill"
                        tempRoll = chara.skills(i_caller)
                    Case "lb_Saves"
                        tempRoll = chara.saves(i_caller)
                End Select



                If s_caller <> "" Then


                    Select Case e.RowIndex
                        Case 0
                            tempRoll.type = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                        Case 1
                            If dg_Rolls.Rows(e.RowIndex).Cells(1).Value <> "" Then
                                tempRoll.name = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                            Else
                                dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                MsgBox("The name field cannot be empty", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                            End If
                        Case 2
                            If dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Contains("/") Then
                                Dim s_ranges As String() = dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Split("/")
                                If s_ranges.Length = 2 And IsNumeric(s_ranges(0)) And IsNumeric(s_ranges(1)) And (s_ranges(0) Mod 1) = 0 And (s_ranges(1) Mod 1) = 0 Then
                                    tempRoll.range = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                                End If
                            Else
                                If dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString = "" Or dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString = Nothing Then
                                    tempRoll.range = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                                Else
                                    dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                    MsgBox("Must have two integer numbers splited by: / ", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                                End If
                            End If
                        Case 3
                            If s_caller = "lb_DCAttacks" Then
                                If dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString = "100" Then
                                    tempRoll.roll = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                                Else
                                    'Dim s_roll_Dc As String() = {""}
                                    'If dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Contains("/") Then
                                    '    s_roll_Dc = dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Split("/")
                                    'End If
                                    'If s_roll_Dc.Length = 3 Then

                                    '    '' verificar si tiene {} etiquetas previamente
                                    '    Dim deleteBool As Boolean = False
                                    '    Dim s_rollDeletedices As String = ""
                                    '    For Each tempChar As Char In s_roll_Dc(0)
                                    '        If tempChar = "{"c Then
                                    '            deleteBool = True
                                    '        ElseIf tempChar = "}"c Then
                                    '            deleteBool = False
                                    '        Else
                                    '            If deleteBool = False Then
                                    '                s_rollDeletedices = s_rollDeletedices + tempChar
                                    '            End If
                                    '        End If
                                    '    Next
                                    '    s_rollDeletedices = s_rollDeletedices.Replace("+", "")
                                    '    s_rollDeletedices = s_rollDeletedices.Replace("-", "")


                                    '    Dim dcFirstMember As String = s_rollDeletedices
                                    '' fin de veriricacion


                                    '' If IsNumeric(dcFirstMember) AndAlso (dcFirstMember Mod 1) = 0 And s_validStats.Contains(s_roll_Dc(1).ToUpper) And s_validDc3rdvar.Contains(s_roll_Dc(2).ToUpper) Then
                                    tempRoll.roll = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                                    ''Else
                                    ''    dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                    ''    MsgBox("The value must be 100 or 3 values split by (/): Number/" + s_validStats + "/" + s_validDc3rdvar, MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                                    ''End If
                                    '' Else
                                    ''dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                    ''MsgBox("The value must be 100 or 3 values split by (/): Number/" + s_validStats + "/" + s_validDc3rdvar, MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                                    '' End If
                                End If
                            Else
                                Dim s_validrol As String = isValidrole(dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Replace(" ", ""))

                                If s_validrol = "" Then
                                    tempRoll.roll = dg_Rolls.Rows(e.RowIndex).Cells(1).Value.ToString.Replace(" ", "")
                                Else
                                    dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                    MsgBox(s_validrol, MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                                End If
                            End If
                        Case 4
                            If IsNumeric(dg_Rolls.Rows(e.RowIndex).Cells(1).Value) Then
                                If (dg_Rolls.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                                    tempRoll.critmultip = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                                End If
                            Else
                                dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                MsgBox("Invalid value: must be an integer")

                            End If
                        Case 5
                            If IsNumeric(dg_Rolls.Rows(e.RowIndex).Cells(1).Value) Then
                                If (dg_Rolls.Rows(e.RowIndex).Cells(1).Value Mod 1) = 0 Then
                                    If dg_Rolls.Rows(e.RowIndex).Cells(1).Value > 0 And dg_Rolls.Rows(e.RowIndex).Cells(1).Value < 21 Then
                                        tempRoll.critrangemin = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                                    Else
                                        dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                        MsgBox("It must be a value between 1 and 20", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                                    End If
                                Else
                                    dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                    MsgBox("It must be a value between 1 and 20", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                                End If
                            Else
                                dg_Rolls.Rows(e.RowIndex).Cells(1).Value = dg_Rolls.Tag
                                MsgBox("It must be a value between 1 and 20", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                            End If
                        Case 6
                            tempRoll.info = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                        Case 7
                            If s_caller = "lb_Heals" Then
                                tempRoll.menuUI = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                            Else
                                tempRoll.futureUse_icon = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                            End If

                        Case 8
                            tempRoll.menuUI = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
                    End Select

                    Select Case s_caller
                        Case "lb_Attacks"
                            chara.attacks(i_caller) = tempRoll
                            lb_Attacks.Items.Item(lb_Attacks.SelectedIndex) = tempRoll.name
                        Case "lb_DCAttacks"
                            chara.attacksDC(i_caller) = tempRoll
                            lb_DCAttacks.Items.Item(lb_DCAttacks.SelectedIndex) = tempRoll.name
                        Case "lb_Heals"
                            chara.healing(i_caller) = tempRoll
                            lb_Heals.Items.Item(lb_Heals.SelectedIndex) = tempRoll.name
                        Case "lb_Skill"
                            chara.skills(i_caller) = tempRoll
                            lb_Skill.Items.Item(lb_Skill.SelectedItems(0).Index).Text = tempRoll.name
                        Case "lb_Saves"
                            chara.saves(i_caller) = tempRoll
                            lb_Saves.Items.Item(lb_Saves.SelectedItems(0).Index).Text = tempRoll.name
                    End Select

                Else
                    'MsgBox("There is no selected list")
                End If

                actualizarJsonText()

            End If

        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub dg_Rolls_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dg_Rolls.CellValidating
        Try
            dg_Rolls.Tag = dg_Rolls.Rows(e.RowIndex).Cells(1).Value
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try

    End Sub


    Private Sub dg_Rolls_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dg_Rolls.DataError
        MsgBox(e.Exception.ToString, MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            actualizarJsonText()
            From_ShowJson.Show()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try

    End Sub


    Private Sub dg_Damages_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dg_Damages.CellValueChanged
        Try

            If avoidmakechangesDamageGrids = True Then
                Exit Sub
            End If

            If e.RowIndex <> -1 Then
                If dg_Damages.Rows(e.RowIndex).Cells(1).Value = Nothing Then
                    dg_Damages.Rows(e.RowIndex).Cells(1).Value = String.Empty
                End If
            End If

            Dim indicedamageindex As Integer = lb_damages.SelectedIndex
            If b_damagedgcanedit = True Then
                Dim b_changename As Boolean = False
                Dim s_caller As String = ""
                Dim l_caller As ListBox = Nothing
                Dim b_doChanges As Boolean = False
                Dim i_caller As Integer = 0
                Dim i_templink As Integer = 0
                Dim templink As Roll = Nothing

                '' ¿Qué listbox me está llamando?
                For Each l_caller In alllists
                    If l_caller.SelectedIndex <> -1 Then
                        s_caller = l_caller.Name
                        i_caller = l_caller.SelectedIndex
                    End If
                Next

                If lb_Skill.SelectedItems.Count > 0 Then
                    s_caller = lb_Skill.Name
                    i_caller = lb_Skill.SelectedItems(0).Index
                End If
                If lb_Saves.SelectedItems.Count > 0 Then
                    s_caller = lb_Saves.Name
                    i_caller = lb_Saves.SelectedItems(0).Index
                End If


                Dim tempRoll As New Roll
                Select Case s_caller
                    Case "lb_Attacks"
                        tempRoll = chara.attacks(i_caller)
                    Case "lb_DCAttacks"
                        tempRoll = chara.attacksDC(i_caller)
                    Case "lb_Heals"
                        tempRoll = chara.healing(i_caller)
                    Case "lb_Skill"
                        tempRoll = chara.skills(i_caller)
                    Case "lb_Saves"
                        tempRoll = chara.saves(i_caller)
                End Select

                templink = tempRoll.link
                If lb_damages.SelectedIndex > 0 Then
                    For index As Integer = 1 To lb_damages.SelectedIndex
                        templink = templink.link
                        i_templink = i_templink + 1
                    Next
                End If
                If s_caller <> "" Then
                    Select Case e.RowIndex
                        Case 0
                            templink.type = dg_Damages.Rows(e.RowIndex).Cells(1).Value
                        Case 1
                            If dg_Damages.Rows(e.RowIndex).Cells(1).Value <> "" And dg_Damages.Rows(e.RowIndex).Cells(1).Value IsNot Nothing Then
                                templink.name = dg_Damages.Rows(e.RowIndex).Cells(1).Value
                                b_changename = True
                            Else
                                avoidmakechangesDamageGrids = True
                                dg_Damages.Rows(e.RowIndex).Cells(1).Value = dg_Damages.Tag
                                avoidmakechangesDamageGrids = False
                                MsgBox("The name field cannot be empty", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                            End If
                        Case 3
                            Dim s_validrol As String = isValidrole(dg_Damages.Rows(e.RowIndex).Cells(1).Value.ToString)
                            'If s_validrol = "" Then
                            templink.roll = dg_Damages.Rows(e.RowIndex).Cells(1).Value
                            'Else
                            '    avoidmakechangesDamageGrids = True
                            '    dg_Damages.Rows(e.RowIndex).Cells(1).Value = dg_Damages.Tag
                            '    avoidmakechangesDamageGrids = False
                            '    MsgBox(s_validrol, MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
                            'End If
                                Case 6
                            templink.info = dg_Damages.Rows(e.RowIndex).Cells(1).Value
                    End Select

                    Select Case s_caller
                        Case "lb_Attacks"
                            chara.attacks(i_caller) = tempRoll
                            'lb_Attacks.Items.Item(lb_Attacks.SelectedIndex) = tempRoll.name //2024/02/05 para evitar un error quizá me arreienta más tarde porque no se por que estaba aquí este código, no parece necesario. Aun así hay que revisar el problema de redundancia al llamar a Datachange sobre si mismo en el grid de daños
                        Case "lb_DCAttacks"
                            chara.attacksDC(i_caller) = tempRoll
                            ''lb_DCAttacks.Items.Item(lb_DCAttacks.SelectedIndex) = tempRoll.name
                        Case "lb_Heals"
                            chara.healing(i_caller) = tempRoll
                            ''lb_Heals.Items.Item(lb_Heals.SelectedIndex) = tempRoll.name
                        Case "lb_Skill"
                            chara.skills(i_caller) = tempRoll
                            ''lb_Skill.Items.Item(lb_Skill.SelectedIndex) = tempRoll.name
                        Case "lb_Saves"
                            chara.saves(i_caller) = tempRoll
                            ''lb_Saves.Items.Item(lb_Saves.SelectedIndex) = tempRoll.name
                    End Select
                    If b_changename = True Then
                        '' regenerar lista
                        lb_damages.Items.Clear()
                        If (lb_Attacks.SelectedIndex <> -1) Then
                            If chara.attacks(lb_Attacks.SelectedIndex).link IsNot Nothing Then
                                Dim tlink As Roll
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
                        Else
                            If chara.attacksDC(lb_DCAttacks.SelectedIndex).link IsNot Nothing Then
                                Dim tlink As Roll
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

                        b_changename = False
                    End If
                    ''  fin regenerar lista.
                Else
                    'MsgBox("There is no selected list")
                End If

            End If
            lb_damages.SelectedIndex = indicedamageindex
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub dg_Damages_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dg_Damages.CellValidating
        Try
            dg_Damages.Tag = dg_Damages.Rows(e.RowIndex).Cells(1).Value
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try

    End Sub

    Private Sub ComboResiCheck(sender As Object, e As EventArgs) Handles cb_vulne_Thunder.CheckedChanged, cb_vulne_Slashing.CheckedChanged, cb_vulne_Radiant.CheckedChanged, cb_vulne_Psychic.CheckedChanged, cb_vulne_Poison.CheckedChanged, cb_vulne_Piercing.CheckedChanged, cb_vulne_Necrotic.CheckedChanged, cb_vulne_Lightning.CheckedChanged, cb_vulne_Force.CheckedChanged, cb_vulne_Fire.CheckedChanged, cb_vulne_Cold.CheckedChanged, cb_vulne_Bludgeoning.CheckedChanged, cb_vulne_Acid.CheckedChanged, cb_resist_Thunder.CheckedChanged, cb_resist_Slashing.CheckedChanged, cb_resist_Radiant.CheckedChanged, cb_resist_Psychic.CheckedChanged, cb_resist_Poison.CheckedChanged, cb_resist_Piercing.CheckedChanged, cb_resist_Necrotic.CheckedChanged, cb_resist_Lightning.CheckedChanged, cb_resist_Force.CheckedChanged, cb_resist_Fire.CheckedChanged, cb_resist_Cold.CheckedChanged, cb_resist_Bludgeoning.CheckedChanged, cb_resist_Acid.CheckedChanged, cb_immu_Thunder.CheckedChanged, cb_immu_Slashing.CheckedChanged, cb_immu_Radiant.CheckedChanged, cb_immu_Psychic.CheckedChanged, cb_immu_Poison.CheckedChanged, cb_immu_Piercing.CheckedChanged, cb_immu_Necrotic.CheckedChanged, cb_immu_Lightning.CheckedChanged, cb_immu_Force.CheckedChanged, cb_immu_Fire.CheckedChanged, cb_immu_Cold.CheckedChanged, cb_immu_Bludgeoning.CheckedChanged, cb_immu_Acid.CheckedChanged, cb_resist_Magic0Bludgeoning.CheckedChanged, cb_resist_Magic0Slashing.CheckedChanged, cb_resist_Magic0Slashing.CheckedChanged, cb_immu_Magic0Bludgeoning.CheckedChanged, cb_immu_Magic0Piercing.CheckedChanged, cb_immu_Magic0Slashing.CheckedChanged, cb_vulne_Magic0Bludgeoning.CheckedChanged, cb_vulne_Magic0Piercing.CheckedChanged, cb_vulne_Magic0Slashing.CheckedChanged, cb_immu_Critical.CheckedChanged
        Try

            If b_checkedchanges = True Then

                Dim senderBox As System.Windows.Forms.CheckBox
                senderBox = sender
                If senderBox.Checked = True Then
                    Dim a_senderBox() As String = senderBox.Name.Split("_")
                    Dim b_reschange As Boolean = True
                    Select Case a_senderBox(1)
                        Case "immu"
                            For Each temp As String In chara.immunity
                                If temp = a_senderBox(2).Replace("0", " ").ToLower Then
                                    b_reschange = False
                                End If
                            Next
                            If b_reschange = True Then
                                chara.immunity.Add(a_senderBox(2).Replace("0", " ").ToLower)
                            End If
                        Case "resist"
                            For Each temp As String In chara.resistance
                                If temp = a_senderBox(2).Replace("0", " ").ToLower Then
                                    b_reschange = False
                                End If
                            Next
                            If b_reschange = True Then
                                chara.resistance.Add(a_senderBox(2).Replace("0", " ").ToLower)
                            End If
                        Case "vulne"
                            For Each temp As String In chara.vulnerability
                                If temp = a_senderBox(2).Replace("0", " ").ToLower Then
                                    b_reschange = False
                                End If
                            Next
                            If b_reschange = True Then
                                chara.vulnerability.Add(a_senderBox(2).Replace("0", " ").ToLower)
                            End If
                    End Select
                End If

                If senderBox.Checked = False Then
                    Dim a_senderBox() As String = senderBox.Name.Split("_")
                    Dim b_reschange As Boolean = False
                    Select Case a_senderBox(1)
                        Case "immu"
                            For Each temp As String In chara.immunity
                                If temp = a_senderBox(2).Replace("0", " ").ToLower Then
                                    b_reschange = True
                                End If
                            Next
                            If b_reschange = True Then
                                chara.immunity.Remove(a_senderBox(2).Replace("0", " ").ToLower)
                            End If
                        Case "resist"
                            For Each temp As String In chara.resistance
                                If temp = a_senderBox(2).Replace("0", " ").ToLower Then
                                    b_reschange = True
                                End If
                            Next
                            If b_reschange = True Then
                                chara.resistance.Remove(a_senderBox(2).Replace("0", " ").ToLower)
                            End If
                        Case "vulne"
                            For Each temp As String In chara.vulnerability
                                If temp = a_senderBox(2).Replace("0", " ").ToLower Then
                                    b_reschange = True
                                End If
                            Next
                            If b_reschange = True Then
                                chara.vulnerability.Remove(a_senderBox(2).Replace("0", " ").ToLower)
                            End If
                    End Select
                End If

            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_skills_add_Click(sender As Object, e As EventArgs) Handles b_skills_add.Click
        Try
            Dim baseroll As New Roll
            baseroll.name = "New"
            baseroll.type = "Public"
            baseroll.roll = "1D20+0"
            chara.skills.Add(baseroll)
            lb_Skill.Items.Clear()
            LoadSkillsList()
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_skills_delete_Click(sender As Object, e As EventArgs) Handles b_skills_delete.Click
        Try
            If lb_Skill.SelectedItems(0).Index <> -1 Then
                chara.skills.RemoveAt(lb_Skill.SelectedItems(0).Index)
                dg_Rolls.Rows.Clear()
                lb_Skill.Items.Clear()

                LoadSkillsList()
            Else
                MsgBox("There is no item selected in the list", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_attacks_add_Click(sender As Object, e As EventArgs) Handles b_attacks_add.Click
        Try
            Dim baseroll As New Roll
            baseroll.name = "New"
            baseroll.type = "Melee"
            baseroll.roll = "1D20"
            baseroll.range = "5/5"
            baseroll.critmultip = "2"
            baseroll.critrangemin = "20"
            If Cmb_AbilityMod.Text <> "" Then
                baseroll.roll = baseroll.roll + "+" + Cmb_AbilityMod.Text
            End If
            If t_attackImport.Text <> "" Then
                baseroll.roll = baseroll.roll + "+" + t_attackImport.Text
            End If
            chara.attacks.Add(baseroll)
            lb_Attacks.Items.Clear()
            For Each roll In chara.attacks
                lb_Attacks.Items.Add(roll.name)
            Next
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_dcattacks_add_Click(sender As Object, e As EventArgs) Handles b_dcattacks_add.Click
        Try
            Dim baseroll As New Roll
            baseroll.name = "New"
            baseroll.type = "Magic"
            baseroll.roll = "10/STR/zero"
            baseroll.range = "5/5"
            'baseroll.critmultip = ""
            'baseroll.critrangemin = ""
            If Cmb_AbilityMod.Text <> "" Then
                baseroll.roll = Cmb_AbilityMod.Text + "+" + baseroll.roll
            End If
            If t_DCImport.Text <> "" Then
                baseroll.roll = t_DCImport.Text + "+" + baseroll.roll
            End If
            If Cmb_AbilityMod.Text <> "" Or t_DCImport.Text <> "" Then
                baseroll.roll = baseroll.roll.Replace("+10", "")
            End If

            chara.attacksDC.Add(baseroll)
            lb_DCAttacks.Items.Clear()
            For Each roll In chara.attacksDC
                lb_DCAttacks.Items.Add(roll.name)
            Next

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_heals_add_Click(sender As Object, e As EventArgs) Handles b_heals_add.Click
        Try
            Dim baseroll As New Roll
            baseroll.name = "New"
            baseroll.type = "Magic"
            baseroll.roll = "1D4+0"
            baseroll.range = "5/5"
            chara.healing.Add(baseroll)
            lb_Heals.Items.Clear()
            For Each roll In chara.healing
                lb_Heals.Items.Add(roll.name)
            Next

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_saves_add_Click(sender As Object, e As EventArgs) Handles b_saves_add.Click
        Try
            Dim baseroll As New Roll
            baseroll.name = "New"
            baseroll.type = "Public"
            baseroll.roll = "1D20+0"
            chara.saves.Add(baseroll)
            lb_Saves.Items.Clear()
            LoadSavesList()

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_attacks_delete_Click(sender As Object, e As EventArgs) Handles b_attacks_delete.Click
        Try
            If lb_Attacks.SelectedIndex <> -1 Then
                chara.attacks.RemoveAt(lb_Attacks.SelectedIndex)
                dg_Rolls.Rows.Clear()
                lb_Attacks.Items.Clear()
                lb_damages.Items.Clear()
                dg_Damages.Rows.Clear()
                For Each roll In chara.attacks
                    lb_Attacks.Items.Add(roll.name)
                Next
            Else
                MsgBox("There is no item selected in the list", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
            End If

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_dcattacks_delete_Click(sender As Object, e As EventArgs) Handles b_dcattacks_delete.Click
        Try
            If lb_DCAttacks.SelectedIndex <> -1 Then
                chara.attacksDC.RemoveAt(lb_DCAttacks.SelectedIndex)
                lb_DCAttacks.Items.Clear()
                dg_Rolls.Rows.Clear()
                lb_damages.Items.Clear()
                dg_Damages.Rows.Clear()

                For Each roll In chara.attacksDC
                    lb_DCAttacks.Items.Add(roll.name)
                Next
            Else
                MsgBox("There is no item selected in the list", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
            End If

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_heals_delete_Click(sender As Object, e As EventArgs) Handles b_heals_delete.Click
        Try
            If lb_Heals.SelectedIndex <> -1 Then
                chara.healing.RemoveAt(lb_Heals.SelectedIndex)
                dg_Rolls.Rows.Clear()
                lb_Heals.Items.Clear()
                For Each roll In chara.healing
                    lb_Heals.Items.Add(roll.name)
                Next
            Else
                MsgBox("There is no item selected in the list")
            End If

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_saves_delete_Click(sender As Object, e As EventArgs) Handles b_saves_delete.Click
        Try
            If lb_Saves.SelectedItems.Count > 0 Then
                chara.saves.RemoveAt(lb_Saves.SelectedItems(0).Index)
                dg_Rolls.Rows.Clear()
                lb_Saves.Items.Clear()
                LoadSavesList()
            Else
                MsgBox("There is no item selected in the list", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
            End If

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_damages_add_Click(sender As Object, e As EventArgs) Handles b_damages_add.Click
        Try
            Dim s_caller As String = ""
            Dim i_caller As Integer = 0
            Dim tempRoll As Roll = Nothing
            Dim baseroll As New Roll
            baseroll.name = "New"
            baseroll.type = ""
            baseroll.roll = "1D4+0"

            For Each l_caller In alllists
                If l_caller.SelectedIndex <> -1 Then
                    s_caller = l_caller.Name
                End If
            Next
            If lb_Skill.SelectedItems.Count > 0 Then
                s_caller = lb_Skill.Name
            End If
            If lb_Saves.SelectedItems.Count > 0 Then
                s_caller = lb_Saves.Name
            End If

            Dim icount As Integer = 0


            Select Case s_caller
                Case "lb_Attacks"

                    Dim templink As Roll = Nothing
                    tempRoll = chara.attacks(lb_Attacks.SelectedIndex)

                    While tempRoll.link IsNot Nothing
                        tempRoll = tempRoll.link
                    End While

                    tempRoll.link = baseroll
                    'chara.attacks(lb_Attacks.SelectedIndex).link = tempRoll

                    lb_damages.Items.Clear()
                    If chara.attacks(lb_Attacks.SelectedIndex).link IsNot Nothing Then
                        Dim tlink As Roll
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
                Case "lb_DCAttacks"
                    Dim templink As Roll = Nothing
                    tempRoll = chara.attacksDC(lb_DCAttacks.SelectedIndex)

                    While tempRoll.link IsNot Nothing
                        tempRoll = tempRoll.link
                    End While

                    tempRoll.link = baseroll
                    'chara.attacks(lb_Attacks.SelectedIndex).link = tempRoll

                    lb_damages.Items.Clear()
                    If chara.attacksDC(lb_DCAttacks.SelectedIndex).link IsNot Nothing Then
                        Dim tlink As Roll
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
            End Select



            ' chara.healing.Add(baseroll)

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_damages_delete_Click(sender As Object, e As EventArgs) Handles b_damages_delete.Click
        Try
            Dim s_caller As String = ""
            Dim i_caller As Integer = 0
            Dim tempRoll As Roll = Nothing

            For Each l_caller In alllists
                If l_caller.SelectedIndex <> -1 Then
                    s_caller = l_caller.Name
                End If
            Next
            If lb_Skill.SelectedItems.Count > 0 Then
                s_caller = lb_Skill.Name
            End If
            If lb_Saves.SelectedItems.Count > 0 Then
                s_caller = lb_Saves.Name
            End If
            If lb_damages.SelectedIndex <> -1 Then
                Select Case s_caller
                    Case "lb_Attacks"
                        Dim i_deleteobject As Integer = 0
                        Dim templink As Roll = Nothing
                        tempRoll = chara.attacks(lb_Attacks.SelectedIndex)

                        While tempRoll.link IsNot Nothing
                            If i_deleteobject <> lb_damages.SelectedIndex Then
                                tempRoll = tempRoll.link
                            Else
                                If tempRoll.link.link IsNot Nothing Then
                                    Dim temp_extra_link As New Roll
                                    temp_extra_link = tempRoll.link.link
                                    tempRoll.link = New Roll(temp_extra_link)
                                Else
                                    tempRoll.link = Nothing
                                End If
                            End If
                            i_deleteobject = i_deleteobject + 1

                        End While
                        lb_damages.Items.Clear()
                        If chara.attacks(lb_Attacks.SelectedIndex).link IsNot Nothing Then
                            Dim tlink As Roll
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
                    Case "lb_DCAttacks"
                        Dim i_deleteobject As Integer = 0
                        Dim templink As Roll = Nothing
                        tempRoll = chara.attacksDC(lb_DCAttacks.SelectedIndex)

                        While tempRoll.link IsNot Nothing
                            If i_deleteobject <> lb_damages.SelectedIndex Then
                                tempRoll = tempRoll.link
                            Else
                                If tempRoll.link.link IsNot Nothing Then
                                    Dim temp_extra_link As New Roll
                                    temp_extra_link = tempRoll.link.link
                                    tempRoll.link = New Roll(temp_extra_link)
                                Else
                                    tempRoll.link = Nothing
                                End If
                            End If
                            i_deleteobject = i_deleteobject + 1

                        End While
                        lb_damages.Items.Clear()
                        If chara.attacksDC(lb_DCAttacks.SelectedIndex).link IsNot Nothing Then
                            Dim tlink As Roll
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
                End Select
                dg_Damages.Rows.Clear()

            Else
                MsgBox("There is no item selected in the list")
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_skills_up_Click(sender As Object, e As EventArgs) Handles b_skills_up.Click
        Try
            If lb_Skill.SelectedItems.Count > 0 Then
                If lb_Skill.SelectedItems(0).Index > 0 Then
                    Dim indi As Integer = lb_Skill.SelectedItems(0).Index - 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.skills(lb_Skill.SelectedItems(0).Index - 1))
                    temproll_u = New Roll(chara.skills(lb_Skill.SelectedItems(0).Index))
                    chara.skills(lb_Skill.SelectedItems(0).Index - 1) = temproll_u
                    chara.skills(lb_Skill.SelectedItems(0).Index) = temproll_d
                    lb_Skill.Items.Clear()

                    LoadSkillsList()
                    lb_Skill.Items(indi).Selected = True
                End If
            End If

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_skills_down_Click(sender As Object, e As EventArgs) Handles b_skills_down.Click
        Try
            If lb_Skill.SelectedItems.Count > 0 Then
                If lb_Skill.SelectedItems(0).Index < lb_Skill.Items.Count - 1 Then
                    Dim indi As Integer = lb_Skill.SelectedItems(0).Index + 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.skills(lb_Skill.SelectedItems(0).Index + 1))
                    temproll_u = New Roll(chara.skills(lb_Skill.SelectedItems(0).Index))
                    chara.skills(lb_Skill.SelectedItems(0).Index + 1) = temproll_u
                    chara.skills(lb_Skill.SelectedItems(0).Index) = temproll_d
                    lb_Skill.Items.Clear()

                    LoadSkillsList()

                    lb_Skill.Items(indi).Selected = True
                End If
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_attacks_up_Click(sender As Object, e As EventArgs) Handles b_attacks_up.Click
        Try
            If lb_Attacks.SelectedIndex <> -1 Then
                If lb_Attacks.SelectedIndex > 0 Then
                    Dim indi As Integer = lb_Attacks.SelectedIndex - 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.attacks(lb_Attacks.SelectedIndex - 1))
                    temproll_u = New Roll(chara.attacks(lb_Attacks.SelectedIndex))
                    chara.attacks(lb_Attacks.SelectedIndex - 1) = temproll_u
                    chara.attacks(lb_Attacks.SelectedIndex) = temproll_d
                    lb_Attacks.Items.Clear()
                    For Each roll In chara.attacks
                        lb_Attacks.Items.Add(roll.name)
                    Next
                    lb_Attacks.SetSelected(indi, True)
                End If
            End If

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_attacksDC_up_Click(sender As Object, e As EventArgs) Handles b_attacksDC_up.Click
        Try
            If lb_DCAttacks.SelectedIndex <> -1 Then
                If lb_DCAttacks.SelectedIndex > 0 Then
                    Dim indi As Integer = lb_DCAttacks.SelectedIndex - 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.attacksDC(lb_DCAttacks.SelectedIndex - 1))
                    temproll_u = New Roll(chara.attacksDC(lb_DCAttacks.SelectedIndex))
                    chara.attacksDC(lb_DCAttacks.SelectedIndex - 1) = temproll_u
                    chara.attacksDC(lb_DCAttacks.SelectedIndex) = temproll_d
                    lb_DCAttacks.Items.Clear()
                    For Each roll In chara.attacksDC
                        lb_DCAttacks.Items.Add(roll.name)
                    Next
                    lb_DCAttacks.SetSelected(indi, True)
                End If
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_heals_up_Click(sender As Object, e As EventArgs) Handles b_heals_up.Click
        Try
            If lb_Heals.SelectedIndex <> -1 Then
                If lb_Heals.SelectedIndex > 0 Then
                    Dim indi As Integer = lb_Heals.SelectedIndex - 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.healing(lb_Heals.SelectedIndex - 1))
                    temproll_u = New Roll(chara.healing(lb_Heals.SelectedIndex))
                    chara.healing(lb_Heals.SelectedIndex - 1) = temproll_u
                    chara.healing(lb_Heals.SelectedIndex) = temproll_d
                    lb_Heals.Items.Clear()
                    For Each roll In chara.healing
                        lb_Heals.Items.Add(roll.name)
                    Next
                    lb_Heals.SetSelected(indi, True)
                End If
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_saves_up_Click(sender As Object, e As EventArgs) Handles b_saves_up.Click
        Try
            If lb_Saves.SelectedItems.Count > 0 Then
                If lb_Saves.SelectedItems(0).Index > 0 Then
                    Dim indi As Integer = lb_Saves.SelectedItems(0).Index - 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.saves(lb_Saves.SelectedItems(0).Index - 1))
                    temproll_u = New Roll(chara.saves(lb_Saves.SelectedItems(0).Index))
                    chara.saves(lb_Saves.SelectedItems(0).Index - 1) = temproll_u
                    chara.saves(lb_Saves.SelectedItems(0).Index) = temproll_d
                    lb_Saves.Items.Clear()

                    LoadSavesList()
                    lb_Saves.Items(indi).Selected = True
                End If
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_attacks_down_Click(sender As Object, e As EventArgs) Handles b_attacks_down.Click
        Try
            If lb_Attacks.SelectedIndex <> -1 Then
                If lb_Attacks.SelectedIndex < lb_Attacks.Items.Count - 1 Then
                    Dim indi As Integer = lb_Attacks.SelectedIndex + 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.attacks(lb_Attacks.SelectedIndex + 1))
                    temproll_u = New Roll(chara.attacks(lb_Attacks.SelectedIndex))
                    chara.attacks(lb_Attacks.SelectedIndex + 1) = temproll_u
                    chara.attacks(lb_Attacks.SelectedIndex) = temproll_d
                    lb_Attacks.Items.Clear()
                    For Each roll In chara.attacks
                        lb_Attacks.Items.Add(roll.name)
                    Next
                    lb_Attacks.SetSelected(indi, True)
                End If
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_attacksdc_down_Click(sender As Object, e As EventArgs) Handles b_attacksdc_down.Click
        Try
            If lb_DCAttacks.SelectedIndex <> -1 Then
                If lb_DCAttacks.SelectedIndex < lb_DCAttacks.Items.Count - 1 Then
                    Dim indi As Integer = lb_DCAttacks.SelectedIndex + 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.attacksDC(lb_DCAttacks.SelectedIndex + 1))
                    temproll_u = New Roll(chara.attacksDC(lb_DCAttacks.SelectedIndex))
                    chara.attacksDC(lb_DCAttacks.SelectedIndex + 1) = temproll_u
                    chara.attacksDC(lb_DCAttacks.SelectedIndex) = temproll_d
                    lb_DCAttacks.Items.Clear()
                    For Each roll In chara.attacksDC
                        lb_DCAttacks.Items.Add(roll.name)
                    Next
                    lb_DCAttacks.SetSelected(indi, True)
                End If
            End If
            actualizarJsonText()

        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_heals_down_Click(sender As Object, e As EventArgs) Handles b_heals_down.Click
        Try
            If lb_Heals.SelectedIndex <> -1 Then
                If lb_Heals.SelectedIndex < lb_Heals.Items.Count - 1 Then
                    Dim indi As Integer = lb_Heals.SelectedIndex + 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.healing(lb_Heals.SelectedIndex + 1))
                    temproll_u = New Roll(chara.healing(lb_Heals.SelectedIndex))
                    chara.healing(lb_Heals.SelectedIndex + 1) = temproll_u
                    chara.healing(lb_Heals.SelectedIndex) = temproll_d
                    lb_Heals.Items.Clear()
                    For Each roll In chara.healing
                        lb_Heals.Items.Add(roll.name)
                    Next
                    lb_Heals.SetSelected(indi, True)
                End If
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_saves_down_Click(sender As Object, e As EventArgs) Handles b_saves_down.Click
        Try
            If lb_Saves.SelectedItems.Count > 0 Then
                If lb_Saves.SelectedItems(0).Index < lb_Saves.Items.Count - 1 Then
                    Dim indi As Integer = lb_Saves.SelectedItems(0).Index + 1
                    Dim temproll_u As New Roll
                    Dim temproll_d As New Roll
                    temproll_d = New Roll(chara.saves(lb_Saves.SelectedItems(0).Index + 1))
                    temproll_u = New Roll(chara.saves(lb_Saves.SelectedItems(0).Index))
                    chara.saves(lb_Saves.SelectedItems(0).Index + 1) = temproll_u
                    chara.saves(lb_Saves.SelectedItems(0).Index) = temproll_d
                    lb_Saves.Items.Clear()

                    LoadSavesList()
                    lb_Saves.Items(indi).Selected = True
                End If
            End If

            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            T_DefaultCPath.Text = My.Settings.characterPath
            T_DefaultAPath.Text = My.Settings.spellPath
            alllists.Add(lb_Attacks)
            alllists.Add(lb_DCAttacks)
            'alllists.Add(lb_Skill)
            alllists.Add(lb_Heals)
            'alllists.Add(lb_Saves)
            ''b_addc.BackColor = Color.LightGray
            b_addat.BackColor = Color.LightGray
            b_adheal.BackColor = Color.LightGray
            Me.Text = Me.Text + " [Ver: " + My.Application.Info.Version.ToString + " ]"
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try


            ' Dim ultrachara As New resistance

            'Dim fichero As String

            If MsgBox("If you continue, the changes in the current file will be lost. Do you want to create a new character?", MsgBoxStyle.YesNo, "Character Editor") = 6 Then

                Dim txtFichero As String = ""
                chara = JsonConvert.DeserializeObject(Of Character)(My.Resources.Jsons.___BaseHeroe)

                cleanall()
                l_creaturename.Text = "New Character"
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
                dg_Principal.Rows.Add("Lv", chara.lv)
                dg_Principal.Rows.Add("Var1", chara.var1)
                dg_Principal.Rows.Add("Var2", chara.var2)
                dg_Principal.Rows.Add("Var3", chara.var3)

                '' Cambio chorizero de color chorra
                If dg_Principal.Rows.Count <> 0 Then
                    changecolors()
                End If

                b_checkedchanges = False

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
                    If im_string.ToLower = "magic bludgeoning" Then
                        cb_resist_Magic0Bludgeoning.Checked = True
                    End If
                    If im_string.ToLower = "magic piercing" Then
                        cb_resist_Magic0Piercing.Checked = True
                    End If
                    If im_string.ToLower = "magic slashing" Then
                        cb_resist_Magic0Slashing.Checked = True
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
                    If im_string.ToLower = "magic bludgeoning" Then
                        cb_immu_Magic0Bludgeoning.Checked = True
                    End If
                    If im_string.ToLower = "magic piercing" Then
                        cb_immu_Magic0Piercing.Checked = True
                    End If
                    If im_string.ToLower = "magic slashing" Then
                        cb_immu_Magic0Slashing.Checked = True
                    End If
                    If im_string.ToLower = "critical" Then
                        cb_immu_Critical.Checked = True
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
                    If im_string.ToLower = "magic bludgeoning" Then
                        cb_vulne_Magic0Bludgeoning.Checked = True
                    End If
                    If im_string.ToLower = "magic piercing" Then
                        cb_vulne_Magic0Piercing.Checked = True
                    End If
                    If im_string.ToLower = "magic slashing" Then
                        cb_vulne_Magic0Slashing.Checked = True
                    End If

                Next

                b_checkedchanges = True

                LoadSavesList()

                LoadSkillsList()

                For Each roll In chara.attacks
                    lb_Attacks.Items.Add(roll.name)
                Next
                For Each roll In chara.attacksDC
                    lb_DCAttacks.Items.Add(roll.name)
                Next
                For Each roll In chara.healing
                    lb_Heals.Items.Add(roll.name)
                Next
            End If
            'Else
            '    MsgBox("The file does not exist", MsgBoxStyle.Exclamation, "Character Editor (RuleSet5EPlugin)")
            'End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        'Try
        '    Dim txtFichero As String = ""
        '    'Dim dlAbrir As New System.Windows.Forms.OpenFileDialog
        '    Dim spell As List(Of Roll) = New List(Of Roll)()
        '    Dim fichero As String
        '    Dim dlAbrir As New System.Windows.Forms.OpenFileDialog
        '    Dim dt_principal As New DataTable
        '    dlAbrir.Filter = "File Dnd5e (*.spell)|*.spell|" & "Todos los archivos (*.*)|*.*"
        '    dlAbrir.Multiselect = False
        '    dlAbrir.CheckFileExists = False
        '    dlAbrir.Title = "Select File"
        '    dlAbrir.ShowDialog()
        '    If dlAbrir.FileName <> "" And System.IO.File.Exists(dlAbrir.FileName) Then
        '        fichero = dlAbrir.FileName
        '        txtFichero = File.ReadAllText(fichero)
        '        If txtFichero <> "" Then
        '            spell = Nothing
        '            spell = JsonConvert.DeserializeObject(Of List(Of Roll))(txtFichero)
        '            For Each sp As Roll In spell
        '                If sp.roll <> "100" Then
        '                    sp.roll = "8" + sp.roll
        '                End If
        '                chara.attacksDC.Add(sp)
        '                lb_DCAttacks.Items.Clear()
        '                For Each roll In chara.attacksDC
        '                    lb_DCAttacks.Items.Add(roll.name)
        '                Next
        '            Next
        '        End If
        '    End If
        '    actualizarJsonText()
        'Catch ex As Exception
        '    MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        'End Try
    End Sub


    Private Sub b_addat_Click(sender As Object, e As EventArgs) Handles b_addat.Click
        Try
            Dim txtFichero As String = ""
            'Dim dlAbrir As New System.Windows.Forms.OpenFileDialog
            Dim spell As List(Of Roll) = New List(Of Roll)()
            Dim fichero As String
            Dim dlAbrirSpell As New System.Windows.Forms.OpenFileDialog
            Dim dt_principal As New DataTable
            dlAbrirSpell.InitialDirectory = T_DefaultAPath.Text
            dlAbrirSpell.Filter = "File Dnd5e (*.spell)|*.spell|" & "Todos los archivos (*.*)|*.*"
            dlAbrirSpell.Multiselect = False
            dlAbrirSpell.CheckFileExists = False
            dlAbrirSpell.Title = "Select File"
            dlAbrirSpell.ShowDialog()
            If dlAbrirSpell.FileName <> "" And System.IO.File.Exists(dlAbrirSpell.FileName) Then
                fichero = dlAbrirSpell.FileName
                txtFichero = File.ReadAllText(fichero)
                If txtFichero <> "" Then
                    spell = Nothing

                    spell = JsonConvert.DeserializeObject(Of List(Of Roll))(txtFichero)
                    For Each sp As Roll In spell
                        If sp.roll.Contains("/") Or sp.roll.Contains("100") Then
                            If Not (sp.roll.Contains("100")) Then
                                Dim temp_roll_import As String
                                temp_roll_import = t_DCImport.Text
                                If Cmb_AbilityMod.Text <> "" Then
                                    If temp_roll_import <> "" Then
                                        temp_roll_import = temp_roll_import + "+" + Cmb_AbilityMod.Text
                                    Else
                                        temp_roll_import = Cmb_AbilityMod.Text
                                    End If

                                End If
                                sp.roll = temp_roll_import + sp.roll
                            End If
                            chara.attacksDC.Add(sp)
                        Else
                            Dim temp_roll_import As String = ""
                            If t_attackImport.Text <> "" Then
                                temp_roll_import = t_attackImport.Text
                            End If
                            If Cmb_AbilityMod.Text <> "" Then
                                temp_roll_import = temp_roll_import + "+" + Cmb_AbilityMod.Text
                            End If

                            If temp_roll_import <> "" Then
                                sp.roll = sp.roll + "+" + temp_roll_import
                            End If
                            chara.attacks.Add(sp)
                        End If
                        lb_Attacks.Items.Clear()
                        lb_DCAttacks.Items.Clear()
                        For Each roll In chara.attacks
                            lb_Attacks.Items.Add(roll.name)
                        Next
                        For Each roll In chara.attacksDC
                            lb_DCAttacks.Items.Add(roll.name)
                        Next
                    Next
                End If
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub l_creaturename_TextChanged(sender As Object, e As EventArgs) Handles l_creaturename.TextChanged
        Try
            If l_creaturename.Text = "" Then
                '' b_addc.Enabled = False
                b_addat.Enabled = False
                b_adheal.Enabled = False
                ''b_addc.BackColor = Color.LightGray
                b_addat.BackColor = Color.LightGray
                b_adheal.BackColor = Color.LightGray
            Else
                ''b_addc.Enabled = True
                b_addat.Enabled = True
                b_adheal.Enabled = True
                ''b_addc.BackColor = Color.Transparent
                b_addat.BackColor = Color.Transparent
                b_adheal.BackColor = Color.Transparent
            End If
            If l_creaturename.Text <> Nothing Then
                characterName = l_creaturename.Text
            End If
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub b_adheal_Click(sender As Object, e As EventArgs) Handles b_adheal.Click
        Try
            Dim txtFichero As String = ""
            'Dim dlAbrir As New System.Windows.Forms.OpenFileDialog
            Dim spell As List(Of Roll) = New List(Of Roll)()
            Dim fichero As String
            Dim dlAbrir As New System.Windows.Forms.OpenFileDialog
            Dim dt_principal As New DataTable
            dlAbrir.InitialDirectory = T_DefaultAPath.Text
            dlAbrir.Filter = "File Dnd5e (*.spell)|*.spell|" & "Todos los archivos (*.*)|*.*"
            dlAbrir.Multiselect = False
            dlAbrir.CheckFileExists = False
            dlAbrir.Title = "Select File"
            dlAbrir.ShowDialog()
            If dlAbrir.FileName <> "" And System.IO.File.Exists(dlAbrir.FileName) Then
                fichero = dlAbrir.FileName
                txtFichero = File.ReadAllText(fichero)
                If txtFichero <> "" Then
                    spell = Nothing

                    spell = JsonConvert.DeserializeObject(Of List(Of Roll))(txtFichero)
                    For Each sp As Roll In spell
                        chara.healing.Add(sp)
                        lb_Heals.Items.Clear()
                        For Each roll In chara.healing
                            lb_Heals.Items.Add(roll.name)
                        Next
                    Next
                End If
            End If
            actualizarJsonText()
        Catch ex As Exception
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim visitGit As New ProcessStartInfo("https://github.com/xerjiomg/Character-Editor-RuleSet5EPlugin-") With {.UseShellExecute = True}
        Process.Start(visitGit)
    End Sub

    Private Sub T_DefaultCPath_TextChanged(sender As Object, e As EventArgs) Handles T_DefaultCPath.TextChanged
        If T_DefaultCPath.Text <> My.Settings.characterPath Then
            My.Settings.characterPath = T_DefaultCPath.Text
        End If
    End Sub

    Private Sub T_DefaultAPath_TextChanged(sender As Object, e As EventArgs) Handles T_DefaultAPath.TextChanged
        If T_DefaultAPath.Text <> My.Settings.spellPath Then
            My.Settings.spellPath = T_DefaultAPath.Text
        End If
    End Sub
    Private Sub c_profSkill_CheckedChanged(sender As Object, e As EventArgs) Handles c_profSkill.CheckedChanged
        If checkProficencyEvent = True Then
            If lb_Skill.SelectedItems.Count > 0 Then
                If c_profSkill.Checked = True Then
                    If Not chara.skills(lb_Skill.SelectedItems(0).Index).roll.Contains("{pb}") Then
                        chara.skills(lb_Skill.SelectedItems(0).Index).roll = chara.skills(lb_Skill.SelectedItems(0).Index).roll + "+{pb}"
                        Dim tempfont As New Font(lb_Skill.Items(lb_Skill.Items.Count - 1).Font, FontStyle.Bold)
                        lb_Skill.Items(lb_Skill.SelectedItems(0).Index).Font = tempfont
                    End If
                Else
                    chara.skills(lb_Skill.SelectedItems(0).Index).roll = chara.skills(lb_Skill.SelectedItems(0).Index).roll.Replace("+{pb}", "")
                    Dim tempfont As New Font(lb_Skill.Items(lb_Skill.Items.Count - 1).Font, FontStyle.Regular)
                    lb_Skill.Items(lb_Skill.SelectedItems(0).Index).Font = tempfont
                End If

                Dim tempindex As Integer = lb_Skill.SelectedItems(0).Index

                lb_Skill.SelectedItems.Clear()
                lb_Skill.SelectedIndices.Add(tempindex)

                actualizarJsonText()
            End If
        End If
    End Sub

    Private Sub c_experSkill_CheckedChanged(sender As Object, e As EventArgs) Handles c_experSkill.CheckedChanged
        If checkProficencyEvent = True Then
            If lb_Skill.SelectedItems.Count > 0 Then
                If c_experSkill.Checked = True Then
                    If Not chara.skills(lb_Skill.SelectedItems(0).Index).roll.Contains("{ex}") Then
                        chara.skills(lb_Skill.SelectedItems(0).Index).roll = chara.skills(lb_Skill.SelectedItems(0).Index).roll + "+{ex}"
                        Dim tempfont As New Font(lb_Skill.Items(lb_Skill.Items.Count - 1).Font, FontStyle.Bold)
                        lb_Skill.Items(lb_Skill.SelectedItems(0).Index).Font = tempfont
                        lb_Skill.Items(lb_Skill.SelectedItems(0).Index).ForeColor = Color.Red
                    End If
                Else
                    chara.skills(lb_Skill.SelectedItems(0).Index).roll = chara.skills(lb_Skill.SelectedItems(0).Index).roll.Replace("+{ex}", "")
                    Dim tempfont As New Font(lb_Skill.Items(lb_Skill.Items.Count - 1).Font, FontStyle.Regular)
                    lb_Skill.Items(lb_Skill.SelectedItems(0).Index).Font = tempfont
                    lb_Skill.Items(lb_Skill.SelectedItems(0).Index).ForeColor = Color.Black
                End If

                Dim tempindex As Integer = lb_Skill.SelectedItems(0).Index

                lb_Skill.SelectedItems.Clear()
                lb_Skill.SelectedIndices.Add(tempindex)

                actualizarJsonText()
            End If
        End If

    End Sub

    Private Sub c_profhalfSkill_CheckedChanged(sender As Object, e As EventArgs) Handles c_profhalfSkill.CheckedChanged
        If checkProficencyEvent = True Then
            If lb_Skill.SelectedItems.Count > 0 Then
                If c_profhalfSkill.Checked = True Then
                    If Not chara.skills(lb_Skill.SelectedItems(0).Index).roll.Contains("{ph}") Then
                        chara.skills(lb_Skill.SelectedItems(0).Index).roll = chara.skills(lb_Skill.SelectedItems(0).Index).roll + "+{ph}"
                        Dim tempfont As New Font(lb_Skill.Items(lb_Skill.Items.Count - 1).Font, FontStyle.Bold)
                        lb_Skill.Items(lb_Skill.SelectedItems(0).Index).Font = tempfont
                        lb_Skill.Items(lb_Skill.SelectedItems(0).Index).ForeColor = Color.Sienna
                    End If
                Else
                    chara.skills(lb_Skill.SelectedItems(0).Index).roll = chara.skills(lb_Skill.SelectedItems(0).Index).roll.Replace("+{ph}", "")
                    Dim tempfont As New Font(lb_Skill.Items(lb_Skill.Items.Count - 1).Font, FontStyle.Regular)
                    lb_Skill.Items(lb_Skill.SelectedItems(0).Index).Font = tempfont
                    lb_Skill.Items(lb_Skill.SelectedItems(0).Index).ForeColor = Color.Black
                End If

                Dim tempindex As Integer = lb_Skill.SelectedItems(0).Index

                lb_Skill.SelectedItems.Clear()
                lb_Skill.SelectedIndices.Add(tempindex)

                actualizarJsonText()
            End If
        End If
    End Sub

    Private Sub c_profSaves_CheckedChanged(sender As Object, e As EventArgs) Handles c_profSaves.CheckedChanged
        If checkProficencyEvent = True Then
            If lb_Saves.SelectedItems.Count > 0 Then
                If c_profSaves.Checked = True Then
                    If Not chara.saves(lb_Saves.SelectedItems(0).Index).roll.Contains("{pb}") Then
                        chara.saves(lb_Saves.SelectedItems(0).Index).roll = chara.saves(lb_Saves.SelectedItems(0).Index).roll + "+{pb}"
                        Dim tempfont As New Font(lb_Saves.Items(lb_Saves.Items.Count - 1).Font, FontStyle.Bold)
                        lb_Saves.Items(lb_Saves.SelectedItems(0).Index).Font = tempfont
                    End If
                Else
                    chara.saves(lb_Saves.SelectedItems(0).Index).roll = chara.saves(lb_Saves.SelectedItems(0).Index).roll.Replace("+{pb}", "")
                    Dim tempfont As New Font(lb_Saves.Items(lb_Saves.Items.Count - 1).Font, FontStyle.Regular)
                    lb_Saves.Items(lb_Saves.SelectedItems(0).Index).Font = tempfont
                End If

                Dim tempindex As Integer = lb_Saves.SelectedItems(0).Index

                lb_Saves.SelectedItems.Clear()
                lb_Saves.SelectedIndices.Add(tempindex)

                actualizarJsonText()
            End If
        End If
    End Sub
End Class
