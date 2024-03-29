﻿Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Character_Editor__RuleSet5EPlugin_.XJ
Imports Character_Editor__RuleSet5EPlugin_.XJ.RuleSet5ECharacterEditor
Imports Newtonsoft.Json

Module FormSupport
    Public characterName As String
    Public s_validDc3rdvar As String = "HALF or ZERO"
    Public s_validStats As String = "STR,CON,DEX,INT,WIS or CHA"
    Public s_validrolls As String = "4,6,8,10,12 or 20"
    Public damagesArray As String() = {"", "acid", "bludgeoning", "cold", "fire", "force", "lightning", "necrotic", "piercing", "poison", "psychic", "radiant", "slashing", "thunder", "magic bludgeoning", "magic piercing", "magic slashing"}
    Public overiteration As Integer = 0
    Public chara As New Character
    Public b_rolldgcanedit As Boolean = True
    Public b_damagedgcanedit As Boolean = True
    Public b_checkedchanges As Boolean = True
    Public alllists As New Collection(Of ListBox)
    Public backcolor As Color = Color.Thistle
    Public avoidmakechangesDamageGrids As Boolean = False
    Public changecolorsList As Boolean = False
    Public checkProficencyEvent As Boolean = True

    'Public m_strMod As Integer = 0
    'Public m_dexMod As Integer = 0
    'Public m_conMod As Integer = 0
    'Public m_intMod As Integer = 0
    'Public m_wisMod As Integer = 0
    'Public m_carMod As Integer = 0
    'Public m_profMod As Integer = 0
    'Public m_MattackMod As Integer = 0
    'Public m_attackMod As Integer = 0
    'Public m_CDMod As Integer = 0


    Function actualizarJsonText()
        Try
            Dim s_outpoutJson As String
            Dim settingjson As New JsonSerializerSettings
            settingjson.NullValueHandling = NullValueHandling.Ignore
            settingjson.DefaultValueHandling = DefaultValueHandling.Ignore
            s_outpoutJson = JsonConvert.SerializeObject(chara, Newtonsoft.Json.Formatting.Indented, settingjson)
            From_ShowJson.s_JsonText.Text = s_outpoutJson
            Return 0


        Catch ex As Exception
            Return 0
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Function


    Function changecolors()
        Try
            '' Cambio chorizero de color chorra
            If Form1.dg_Principal.Rows.Count <> 0 Then
                If Form1.dg_Principal.Rows(0).Cells(1).Value.ToString.ToLower = "true" Then
                    backcolor = Color.Linen
                ElseIf Form1.dg_Principal.Rows(0).Cells(1).Value.ToString.ToLower = "false" Then
                    backcolor = Color.Lavender
                Else
                    backcolor = Color.Coral
                End If
                Form1.BackColor = backcolor
                Form1.panel_lists.BackColor = backcolor
                Form1.panel_resi.BackColor = backcolor
                Form1.panel_rolls.BackColor = backcolor
                Form1.Panelbase.BackColor = backcolor

                Form1.dg_Damages.Columns(0).DefaultCellStyle.BackColor = backcolor
                Form1.dg_Principal.Columns(0).DefaultCellStyle.BackColor = backcolor
                Form1.dg_Rolls.Columns(0).DefaultCellStyle.BackColor = backcolor
                'Form1.dg_Principal.ColumnHeadersDefaultCellStyle.BackColor = backcolor  ''MODO ZEBRA
                'Form1.dg_Rolls.ColumnHeadersDefaultCellStyle.
            End If
            Return 0

        Catch ex As Exception
            Return 0
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Function
    Function isValidrole(s_rolldices As String) As String
        Try
            Dim deleteBool As Boolean = False
            Dim s_rollDeletedices As String = ""
            For Each tempChar As Char In s_rolldices
                If tempChar = "{"c Then
                    deleteBool = True
                ElseIf tempChar = "}"c Then
                    deleteBool = False
                Else
                    If deleteBool = False Then
                        s_rollDeletedices = s_rollDeletedices + tempChar
                    End If
                End If
            Next
            s_rolldices = s_rollDeletedices
            s_rolldices = s_rolldices.ToUpper
            If s_rolldices = Nothing Or s_rolldices.Trim = "" Then
                Return "The roll field cannot be empty"
            End If

            Dim n_rolldices As String = s_rolldices.Replace("+", "").Replace("D", "").Replace("-", "").Replace(" ", "")

            If IsNumeric(n_rolldices) Then
                If (n_rolldices Mod 1) = 0 Then
                    If s_rolldices.Contains("D") Then
                        Dim s_rolvalues() As String = s_rolldices.Split("+"c, "-"c)
                        For Each s_rolvalue As String In s_rolvalues
                            If s_rolvalue.Contains("D") Then
                                Dim s_s_rolvalues() As String = s_rolvalue.Split("D")
                                If s_s_rolvalues(0) <> "" Then
                                    If s_validrolls.Contains(s_s_rolvalues(1)) And s_s_rolvalues(1) <> "" Then
                                        Return ""
                                    Else
                                        Return "The number of faces of the die must be valid: " + s_validrolls + " | " + s_rolvalue.ToString + " > Not valid"
                                    End If
                                Else
                                    Return "Enter the amount of dice to roll > ?d" + s_validrolls.Contains(s_s_rolvalues(1)).ToString
                                End If
                            End If
                        Next
                    End If
                Else
                    Return "Only these values are allowed: numbers and d,D,+,-"
                End If
            Else
                Return "Only these values are allowed: numbers and d,D,+,-"
            End If

            Return ""
        Catch ex As Exception
            Return ""
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Function

    Function cleanall()
        Try
            For Each deletelist As ListBox In Form1.panel_resi.Controls.OfType(Of ListBox)
                deletelist.Items.Clear()
            Next
            For Each deletelist As ListBox In Form1.panel_rolls.Controls.OfType(Of ListBox)
                deletelist.Items.Clear()
            Next
            For Each deletelist As ListBox In alllists
                deletelist.Items.Clear()
            Next
            Form1.lb_Skill.Clear()
            Form1.lb_Saves.Clear()
            Form1.dg_Damages.Rows.Clear()
            Form1.dg_Rolls.Rows.Clear()
            Form1.dg_Principal.Rows.Clear()
            b_checkedchanges = False
            For Each a As System.Windows.Forms.CheckBox In Form1.panel_resi.Controls.OfType(Of System.Windows.Forms.CheckBox)
                a.Checked = False
            Next
            b_checkedchanges = True


            Return 0
        Catch ex As Exception
            Return 0
            MsgBox("Unexpected Error:it is recommended to save your work and restart the application." + Chr(10) + Chr(10) + "Error: [" + ex.Message + "]", MsgBoxStyle.Critical, "Character Editor (RuleSet5EPlugin)")
        End Try
    End Function
    Sub LoadSkillsList()
        For Each roll In chara.skills
            Form1.lb_Skill.Items.Add(roll.name)
            If chara.skills(Form1.lb_Skill.Items.Count - 1).roll.Contains("{pb}") Then
                Dim tempfont As New Font(Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font, FontStyle.Bold)
                Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font = tempfont
            End If
            If chara.skills(Form1.lb_Skill.Items.Count - 1).roll.Contains("{ex}") Then
                'Dim tempfont As New Font(Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font, FontStyle.Regular)
                'Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font = tempfont
                Dim tempfont As New Font(Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font, FontStyle.Bold)
                Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font = tempfont
                Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).ForeColor = Color.Red
            End If
            If chara.skills(Form1.lb_Skill.Items.Count - 1).roll.Contains("{ph}") Then
                'Dim tempfont As New Font(Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font, FontStyle.Bold)
                'Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font = tempfont
                Dim tempfont As New Font(Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font, FontStyle.Bold)
                Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).Font = tempfont
                Form1.lb_Skill.Items(Form1.lb_Skill.Items.Count - 1).ForeColor = Color.Sienna
            End If
        Next
    End Sub
    Sub LoadSavesList()
        For Each roll In chara.saves
            Form1.lb_Saves.Items.Add(roll.name)
            If chara.saves(Form1.lb_Saves.Items.Count - 1).roll.Contains("{pb}") Then
                Dim tempfont As New Font(Form1.lb_Saves.Items(Form1.lb_Saves.Items.Count - 1).Font, FontStyle.Bold)
                Form1.lb_Saves.Items(Form1.lb_Saves.Items.Count - 1).Font = tempfont
            End If
        Next
    End Sub
End Module
