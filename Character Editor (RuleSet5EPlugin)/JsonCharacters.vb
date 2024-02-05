Imports System.ComponentModel
Imports Newtonsoft.Json

Namespace XJ
    Partial Public Class RuleSet5ECharacterEditor
        Public Class Character
            <DefaultValue("")>
            Public Property NPC As Boolean = False
            Public Property reach As Integer = 5
            Public Property ac As String = "8"
            Public Property hp As String = "10"
            Public Property str As String = "10"
            Public Property dex As String = "10"
            Public Property con As String = "10"
            Public Property Int As String = "10"
            Public Property wis As String = "10"
            Public Property cha As String = "10"
            Public Property speed As String = ""
            <DefaultValue("")>
            Public Property lv As String = ""
            <DefaultValue("")>
            Public Property var1 As String = ""
            <DefaultValue("")>
            Public Property var2 As String = ""
            <DefaultValue("")>
            Public Property var3 As String = ""

            Public Property attacks As List(Of Roll) = New List(Of Roll)()
            Public Property attacksDC As List(Of Roll) = New List(Of Roll)()
            Public Property saves As List(Of Roll) = New List(Of Roll)()
            Public Property skills As List(Of Roll) = New List(Of Roll)()
            Public Property healing As List(Of Roll) = New List(Of Roll)()
            Public Property resistance As List(Of String) = New List(Of String)()
            Public Property vulnerability As List(Of String) = New List(Of String)()
            Public Property immunity As List(Of String) = New List(Of String)()


        End Class

        Public Class Roll
            Public Property name As String = ""
            Public Property type As String = ""
            Public Property roll As String = "100"
            <DefaultValue("")>
            Public Property critrangemin As String = ""
            <DefaultValue("")>
            Public Property critmultip As String = ""
            <DefaultValue("")>
            Public Property range As String = ""
            <DefaultValue("")>
            Public Property info As String = ""
            <DefaultValue("")>
            Public Property futureUse_icon As String = ""
            Public Property link As Roll = Nothing

            Public Sub New()

            End Sub

            Public Sub New(ByVal source As Roll)
                name = source.name
                type = source.type
                roll = source.roll
                critrangemin = source.critrangemin
                critmultip = source.critmultip
                range = source.range
                info = source.info
                futureUse_icon = source.futureUse_icon
                If source.link Is Nothing Then
                    link = Nothing
                Else
                    link = New Roll(source.link)
                End If
            End Sub
        End Class
    End Class
End Namespace


