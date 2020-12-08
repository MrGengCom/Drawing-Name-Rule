Option Explicit On
Imports System
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Middle.Services.Hidden
Imports Ingr.SP3D.Systems.Middle
Imports System.Collections.ObjectModel
Imports System.IO


''' <summary>
''' Auther:  Geng Lining
''' Date created:  Jul 27, 2020
''' 
''' History: 
''' Version 000 : 最初版
''' </summary>
''' 
''' <remarks>
''' ISO Drawing命名规则示例代码，自带的源码是VB6，我改了一版.NET的，大家可以在此基础上自行扩展修改
''' </remarks>
Public Class DBQ
    Inherits NameRuleBase
    Public Overrides Sub ComputeName(ByVal oObject As BusinessObject,
            ByVal oParents As ReadOnlyCollection(Of BusinessObject),
            ByVal oActiveEntity As BusinessObject)

        Try
            '获取DBQ Drawing中的过滤对象
            Dim oSheet As BusinessObject = oObject
            Dim sSheetName As String = ""
            'ISO图纸通过下面关系直接取到过滤进去的对象，如Line，Run，Spool等
            '如果是Ortho图纸，获取对象的方法会不一样，需要先跳转到View，再从View到模型对象
            Dim colBo As ReadOnlyCollection(Of BusinessObject) = oSheet.GetRelationship("SheetToDrawingTarget", "DrawingTarget").TargetObjects
            'Dim rDrawingToSheet As RelationCollection = oSheet.GetRelationship("ObjectHasOutput", "ObjectHasOutput_Dest")
            'Dim oDoc = rDrawingToSheet.TargetObjects.Item(0)

            If TypeOf colBo.Item(0) Is Pipeline Then

                Dim oPipeline As Pipeline = colBo.Item(0)
                Dim strPipeline As String = oPipeline.Name
                Dim oPipSys As PipingSystem = oPipeline.SystemParent
                Dim oSys1 As BusinessObject = oPipSys.SystemParent
                Dim strSys1 As String = oSys1.GetPropertyValue("IJNamedItem", "Name").ToString
                If TypeOf oSys1 Is ISystemChild Then
                    Dim oSysChild As ISystemChild = TryCast(oSys1, ISystemChild)
                    Dim oSys2 As BusinessObject = oSysChild.SystemParent
                    Dim strSys2 As String = oSys2.GetPropertyValue("IJNamedItem", "Name").ToString
                    sSheetName = strSys2 & "-" & strSys1 & "-" & strPipeline
                Else
                    sSheetName = strSys1 & "-" & strPipeline
                End If
            End If

            oSheet.SetPropertyValue(sSheetName, "IJNamedItem", "Name")
            

        Catch ex As Exception
            If Not Err.Source.Equals("NotifyUser") Then
                Throw New Exception("Unexpected error:  DrawingNameRules,DrawingNameRules.DBQ")
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Get the parent object of the system being named.
    ''' </summary>
    ''' <param name="oEntity">Input.  System object whose name is being checked.</param>
    ''' <returns>Nothing</returns>
    ''' <remarks></remarks>
    Public Overrides Function GetNamingParents(ByVal oEntity As BusinessObject) As Collection(Of BusinessObject)

        GetNamingParents = Nothing

    End Function

    ''' <summary>
    ''' Add an error to the errors collection that indicates a unacceptable name has been
    ''' corrected.
    ''' </summary>
    ''' <param name="iErrNum">Input.  Error number that indicates either a blank name or
    ''' a duplicate name has been corrected.</param>
    ''' <param name="strSource">Input.  Name of calling subroutine.</param>
    ''' <param name="strDescription">Input.  Description of the problem.</param>
    ''' <param name="strContext">Input.  Error context.</param>
    ''' <remarks>
    ''' Input parameter strContext should be set to "NAMING" for callers in SystemsAndSpecs
    ''' to be able to handle this correctly - by notifying the user but not raising
    ''' the error any higher.
    ''' </remarks>
    Private Sub NotifyUser(ByVal iErrNum As Integer, ByVal strSource As String,
            ByVal strDescription As String, ByVal strContext As String)

        ' On initial creation we are waiting for a logging service to be provided by the CommonApp
        ' team.  See DI CP122309 - Create a logger service.

    End Sub

    ''' <summary>
    ''' Get a collection of all the system children of a parent system.
    ''' </summary>
    ''' <param name="oParent">Input.  Parent system.</param>
    ''' <returns>Collection of the system children.</returns>
    ''' <remarks></remarks>
    Private Function GetSystemChildren(ByVal oParent As BusinessObject) As ReadOnlyCollection(Of BusinessObject)
        GetSystemChildren = Nothing
        Try

            ' Get relation collection for "SystemHierarchy" relationship with "SystemChildren" rolename
            Dim oRelationCollection As RelationCollection
            oRelationCollection = oParent.GetRelationship("SystemHierarchy", "SystemChildren")

            If Not oRelationCollection Is Nothing Then
                GetSystemChildren = oRelationCollection.TargetObjects
            End If

        Catch ex As Exception
            Throw New Exception("Unexpected error:  SystemNameRulesNetVB.UserDefinedNameRule.GetSystemChildren")
        End Try

    End Function

End Class

