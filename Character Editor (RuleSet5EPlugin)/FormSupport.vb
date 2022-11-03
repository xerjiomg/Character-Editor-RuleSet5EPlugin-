Imports System.IO
Imports System.Text
Imports Character_Editor__RuleSet5EPlugin_.XJ
Imports Character_Editor__RuleSet5EPlugin_.XJ.RuleSet5ECharacterEditor

Module FormSupport
    Public s_validDc3rdvar As String = "HALF or ZERO"
    Public s_validStats As String = "STR,CON,DEX,INT,WIS or CHA"
    Public s_validrolls As String = "4,6,8,10,12 or 20"
    Public damagesArray As String() = {"acid", "bludgeoning", "cold", "fire", "force", "lightning", "necrotic", "piercing", "poison", "psychic", "radiant", "slashing", "thunder"}
    Public overiteration As Integer = 0
    Public chara As New Character


    Function isValidrole(s_rolldices As String) As String

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
                                If s_validrolls.Contains(s_s_rolvalues(1)) Then
                                    Return ""
                                Else
                                    Return "The number of faces of the die must be valid:" + s_rolvalue.ToString
                                End If
                            Else
                                Return "Enter the amount of dice to roll > ?d" + s_validrolls.Contains(s_s_rolvalues(1)).ToString
                            End If
                        End If
                    Next
                End If
            Else
                Return "Only these values are allowed: numbers and d,D,+,- "
            End If
        Else
            Return "Only these values are allowed: numbers and d,D,+,- "
        End If

        Return False
    End Function

    Function cleanall()
        For Each deletelist As ListBox In Form1.panel_resi.Controls.OfType(Of ListBox)
            deletelist.Items.Clear()
        Next
        For Each deletelist As ListBox In Form1.panel_rolls.Controls.OfType(Of ListBox)
            deletelist.Items.Clear()
        Next
        For Each deletelist As ListBox In Form1.panel_damages.Controls.OfType(Of ListBox)
            deletelist.Items.Clear()
        Next
        For Each deletelist As ListBox In Form1.panel_lists.Controls.OfType(Of ListBox)
            deletelist.Items.Clear()
        Next
        Form1.dg_Damages.Rows.Clear()
        Form1.dg_Rolls.Rows.Clear()
        Form1.dg_Principal.Rows.Clear()

        Form1.cb_resist_Bludgeoning.Checked = False
        Form1.cb_resist_Piercing.Checked = False
        Form1.cb_resist_Slashing.Checked = False
        Form1.cb_resist_Acid.Checked = False
        Form1.cb_resist_Cold.Checked = False
        Form1.cb_resist_Fire.Checked = False
        Form1.cb_resist_Force.Checked = False
        Form1.cb_resist_Lightning.Checked = False
        Form1.cb_resist_Necrotic.Checked = False
        Form1.cb_resist_Poison.Checked = False
        Form1.cb_resist_Radiant.Checked = False
        Form1.cb_resist_Psychic.Checked = False
        Form1.cb_resist_Thunder.Checked = False
        Form1.cb_immu_Bludgeoning.Checked = False
        Form1.cb_immu_Piercing.Checked = False
        Form1.cb_immu_Slashing.Checked = False
        Form1.cb_immu_Acid.Checked = False
        Form1.cb_immu_Cold.Checked = False
        Form1.cb_immu_Fire.Checked = False
        Form1.cb_immu_Force.Checked = False
        Form1.cb_immu_Lightning.Checked = False
        Form1.cb_immu_Necrotic.Checked = False
        Form1.cb_immu_Poison.Checked = False
        Form1.cb_immu_Radiant.Checked = False
        Form1.cb_immu_Psychic.Checked = False
        Form1.cb_immu_Thunder.Checked = False
        Form1.cb_vulne_Bludgeoning.Checked = False
        Form1.cb_vulne_Piercing.Checked = False
        Form1.cb_vulne_Slashing.Checked = False
        Form1.cb_vulne_Acid.Checked = False
        Form1.cb_vulne_Cold.Checked = False
        Form1.cb_vulne_Fire.Checked = False
        Form1.cb_vulne_Force.Checked = False
        Form1.cb_vulne_Lightning.Checked = False
        Form1.cb_vulne_Necrotic.Checked = False
        Form1.cb_vulne_Poison.Checked = False
        Form1.cb_vulne_Radiant.Checked = False
        Form1.cb_vulne_Psychic.Checked = False
        Form1.cb_vulne_Thunder.Checked = False
        Return 0
    End Function
End Module
