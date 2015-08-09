#Region "COPL"
#If False Then
AutoUpdate is a Derivative Work of an application of the same name.
The Author of the original Work is Eduardo Oliveira.
The Work can be found at the following URL:
http://www.codeproject.com/Articles/16453/Application-Auto-Update-Revisited
The Work is Copyright (c) Eduardo Oliveira 2006 under
"The Code Project Open License (CPOL)."

According to said license (of which a copy is supplied in the COPL file),
"You may otherwise modify Your copy of this Work (excluding the Articles)
in any way to create a Derivative Work,
provided that You insert a prominent
notice in each changed file stating how,
when and where You changed that file."

I, Russell Dillin, changed the file on 08/09/2015 to simplify it for my own
purposes as well as to update it to use .NET Framework 4.5.
Deprecated resources were exchanged for Microsoft.VisualBasic.FileIO. The
file has been extensively editted in order to keep the line length under
80 columns in length. Additionally, simplified the code by removing the
parser-syntax functionality and changing it to use files on the local
network instead of using any remote network protocols.
#End If
#End Region

Imports Microsoft.VisualBasic.FileIO
Imports System.Reflection

Public Module AutoUpdate
#Region "Properties"
	Private m_RootPath As String = String.Empty
	Private m_UpdateFile As String = "_update.dat"
	Private m_ErrorMessage As String = "ERROR IN AUTO-UPDATE!"

	Public Property RootPath As String
		Get
			Return m_RootPath
		End Get
		Set(ByVal value As String)
			m_RootPath = value
		End Set
	End Property

	Public Property UpdateFile As String
		Get
			Return m_UpdateFile
		End Get
		Set(ByVal value As String)
			m_UpdateFile = value
		End Set
	End Property

	Public Property ErrorMessage As String
		Get
			Return m_ErrorMessage
		End Get
		Set(ByVal value As String)
			m_ErrorMessage = value
		End Set
	End Property

#End Region
#Region "GetVersion"
	Private Function GetVersion(ByVal Version As String) _
								As String
		Dim x() As String = Split(Version, ".")
		Return String.Format("{0:00}{1:00}{2:00}{3:00}", _
							 Int(x(0)), _
							 Int(x(1)), _
							 Int(x(2)), _
							 Int(x(3)))
	End Function
#End Region

	''' <summary>
	''' Updates executable file. 
	''' </summary>
	''' <param name="path">Root Path</param>
	''' <returns>
	''' True on update/error,
	''' False on up-to-date
	''' </returns>
	''' <remarks></remarks>
	Public Function UpdateEXE(Optional ByVal path _
							  As String = "") _
							  As Boolean
		If path = String.Empty Then
			path = RootPath
		Else
			RootPath = path
		End If
		Dim update As String = UpdateFile

		Dim result As Boolean = False
		Dim dirSep As String = "\"
		Dim deleteExt As String = ".del"
		Dim executableExt As String = ".exe"
		Dim assemblyName As String = Assembly _
									 .GetEntryAssembly _
									 .GetName _
									 .Name
		Try
			'Delete old files, if exist.
			Dim dels As IReadOnlyCollection(Of String) = _
				FileSystem.GetFiles(Application.StartupPath, _
									SearchOption.SearchAllSubDirectories, _
									"*" & deleteExt)
			For Each fileToDelete As String In dels
				FileSystem.DeleteFile(fileToDelete, _
									  UIOption.OnlyErrorDialogs, _
									  RecycleOption.DeletePermanently, _
									  UICancelOption.ThrowException)
			Next
			'Get the update file content.
			Dim currVer As String = New String(String.Empty)
			Dim currVerPath As String = New String(String.Empty)
			Dim searchPath As String = path & dirSep & update
			If FileSystem.FileExists(searchPath) Then
				currVer = FileSystem _
					.OpenTextFileReader(searchPath) _
					.ReadLine _
					.Trim
				searchPath = path & dirSep & GetVersion(currVer)
				If FileSystem.DirectoryExists(searchPath) Then
					currVerPath = searchPath & dirSep
				End If
				'If currVer is greater than Version of running Application,
				Dim appName As String = New String(String.Empty)
				Dim startupPath As String = Application.StartupPath & dirSep
				If Integer.Parse(GetVersion(currVer)) > _
					Integer.Parse(GetVersion(Application.ProductVersion)) Then
					appName = Application.ProductName & _
							  "-" & GetVersion(currVer) & executableExt
					'Copy new version to currenct location
					FileSystem.CopyFile(currVerPath & appName, _
										startupPath & appName)
					'Rename running version
					FileSystem.RenameFile(Application.ExecutablePath, _
										  Now _
										  .TimeOfDay _
										  .TotalMilliseconds & deleteExt)
					'Rename new version
					FileSystem.RenameFile(startupPath & appName, _
										  Application.ProductName & _
										  executableExt)
					result = True
				End If
			End If
		Catch ex As Exception
			result = True
			MsgBox(m_ErrorMessage & vbCr & _
				   ex.Message & vbCr & _
				   "Assembly: " & assemblyName, _
				   MsgBoxStyle.Critical, _
				   Application.ProductName)
		End Try
		Return result
	End Function
End Module
