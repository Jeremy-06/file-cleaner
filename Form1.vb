Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Public Class Form1
    Private fileList As New List(Of FileInfo)
    Private selectedFile As FileInfo

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeForm()
        LoadFileCategories()
    End Sub

    Private Sub TryOpenExternal(filePath As String)
        Try
            If String.IsNullOrWhiteSpace(filePath) Then Return

            Dim psi As New ProcessStartInfo()
            psi.FileName = filePath
            psi.UseShellExecute = True
            psi.ErrorDialog = False

            Process.Start(psi)
        Catch ex As System.ComponentModel.Win32Exception
            ' Error code 1223 (0x4C7) is "The operation was canceled by the user"
            ' Silently ignore this specific case to avoid alarming the user
            Const ERROR_CANCELLED As Integer = 1223
            If ex.NativeErrorCode <> ERROR_CANCELLED Then
                MessageBox.Show($"Cannot open file: {ex.Message}", "Open Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Cannot open file: {ex.Message}", "Open Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    Private Sub InitializeForm()
        ' Setup DataGridView columns
        With DataGridView1
            .Columns.Clear()
            .Columns.Add("Name", "File Name")
            .Columns.Add("Size", "Size (MB)")
            .Columns.Add("Modified", "Date Modified")
            .Columns.Add("Path", "Full Path")
            .Columns("Size").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        End With

        ' Setup TreeView
        With TreeView1
            .Nodes.Clear()
            .ShowLines = True
            .ShowPlusMinus = True
        End With

        ' Setup Panel for preview
        Panel1.BackColor = Color.LightGray
        Panel1.BorderStyle = BorderStyle.FixedSingle
    End Sub

    Private Sub LoadFileCategories()
        TreeView1.Nodes.Clear()

        ' Add main categories
        Dim rootNode As TreeNode = TreeView1.Nodes.Add("File Categories")
        rootNode.Nodes.Add("Images", "Images (.jpg, .png, .gif, .bmp)")
        rootNode.Nodes.Add("Documents", "Documents (.pdf, .docx, .txt, .xlsx)")
        rootNode.Nodes.Add("Videos", "Videos (.mp4, .avi, .mkv, .mov)")
        rootNode.Nodes.Add("Audio", "Audio (.mp3, .wav, .flac, .m4a)")
        rootNode.Nodes.Add("Archives", "Archives (.zip, .rar, .7z)")
        rootNode.Nodes.Add("Executables", "Programs (.exe, .msi)")
        rootNode.Nodes.Add("Temporary", "Temporary Files (.tmp, .temp)")
        rootNode.Nodes.Add("Large", "Large Files (>100MB)")
        rootNode.Nodes.Add("Old", "Old Files (>1 year)")

        rootNode.Expand()
    End Sub

    'KEEP
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim index As Integer = DataGridView1.SelectedRows(0).Index
            DataGridView1.Rows.RemoveAt(index)
            MessageBox.Show("File marked as KEEP and removed from cleanup list.", "Keep File", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Please select a file to keep.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    'DELETE
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim filePath As String = DataGridView1.SelectedRows(0).Cells("Path").Value.ToString()
            Dim fileName As String = DataGridView1.SelectedRows(0).Cells("Name").Value.ToString()

            ' Block deletions inside protected locations (system/dev folders) or of protected project files
            If IsProtectedPath(filePath) OrElse IsProtectedFileType(filePath) Then
                MessageBox.Show("This file is in a protected system/project folder and will not be deleted.",
                              "Protected Location", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim result As DialogResult = MessageBox.Show($"Are you sure you want to delete '{fileName}'?",
                                                        "Confirm Delete",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Try
                    ' Clear preview first to release any file handles
                    Panel1.Controls.Clear()
                    GC.Collect() ' Force garbage collection
                    GC.WaitForPendingFinalizers()

                    ' Check if file is in use and try to delete
                    If IsFileInUse(filePath) Then
                        Dim retryResult As DialogResult = MessageBox.Show(
                            $"The file '{fileName}' is currently in use by another program." & vbCrLf & vbCrLf &
                            "Would you like to:" & vbCrLf &
                            "• Yes - Close preview and try again" & vbCrLf &
                            "• No - Skip this file" & vbCrLf &
                            "• Cancel - Stop operation",
                            "File In Use",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning)

                        Select Case retryResult
                            Case DialogResult.Yes
                                ' Try to force close handles and retry
                                System.Threading.Thread.Sleep(1000)
                                GC.Collect()
                                GC.WaitForPendingFinalizers()

                                If IsFileInUse(filePath) Then
                                    MessageBox.Show($"File '{fileName}' is still in use. Please close any programs using this file and try again.",
                                                  "Still In Use", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                    Return
                                End If
                            Case DialogResult.No
                                Return ' Skip this file
                            Case DialogResult.Cancel
                                Return ' Cancel operation
                        End Select
                    End If

                    ' Move to Recycle Bin instead of permanent delete
                    FileSystem.DeleteFile(filePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin)

                    ' Remove from DataGridView
                    Dim index As Integer = DataGridView1.SelectedRows(0).Index
                    DataGridView1.Rows.RemoveAt(index)

                    MessageBox.Show($"'{fileName}' has been moved to Recycle Bin.", "Delete Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch unauthorizedEx As UnauthorizedAccessException
                    MessageBox.Show($"Access denied. The file '{fileName}' may be:" & vbCrLf &
                                  "• In use by another program" & vbCrLf &
                                  "• Read-only or protected" & vbCrLf &
                                  "• Located in a system folder" & vbCrLf & vbCrLf &
                                  "Please check file permissions and try again.",
                                  "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Catch ioEx As IOException
                    MessageBox.Show($"Cannot delete '{fileName}' because:" & vbCrLf &
                                  "• The file is currently open in another program" & vbCrLf &
                                  "• Another process is using this file" & vbCrLf & vbCrLf &
                                  "Please close any programs using this file and try again.",
                                  "File In Use", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Catch ex As Exception
                    MessageBox.Show($"Error deleting file: {ex.Message}", "Delete Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select a file to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Function IsFileInUse(filePath As String) As Boolean
        Try
            Using fs As FileStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                Return False ' File is not in use
            End Using
        Catch
            Return True ' File is in use
        End Try
    End Function

    'REFRESH
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TreeView1.SelectedNode IsNot Nothing AndAlso TreeView1.SelectedNode.Name <> "" Then
            LoadFilesForCategory(TreeView1.SelectedNode.Name)
        Else
            MessageBox.Show("Please select a file category first.", "No Category Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        If e.Node.Name <> "" Then
            LoadFilesForCategory(e.Node.Name)
        End If
    End Sub

    Private Sub LoadFilesForCategory(category As String)
        DataGridView1.Rows.Clear()
        Panel1.Controls.Clear()

        Try
            ' Show loading message
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()

            Dim files As List(Of FileInfo) = ScanFilesForCategory(category)

            For Each file As FileInfo In files
                Try
                    Dim sizeInMB As Double = Math.Round(file.Length / (1024 * 1024), 2)
                    DataGridView1.Rows.Add(file.Name, sizeInMB, file.LastWriteTime.ToString("yyyy-MM-dd HH:mm"), file.FullName)
                Catch
                    ' Skip files that can't be accessed
                    Continue For
                End Try
            Next

            ' Sort by size (largest first)
            DataGridView1.Sort(DataGridView1.Columns("Size"), System.ComponentModel.ListSortDirection.Descending)

        Catch ex As Exception
            MessageBox.Show($"Error loading files: {ex.Message}", "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function ScanFilesForCategory(category As String) As List(Of FileInfo)
        Dim files As New List(Of FileInfo)
        Dim extensions As String() = {}
        Dim scanAllDrives As Boolean = True

        ' Define extensions for each category
        Select Case category.ToLower()
            Case "images"
                extensions = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"}
            Case "documents"
                extensions = {".pdf", ".doc", ".docx", ".txt", ".rtf", ".xlsx", ".xls", ".ppt", ".pptx"}
            Case "videos"
                extensions = {".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".m4v"}
            Case "audio"
                extensions = {".mp3", ".wav", ".flac", ".aac", ".m4a", ".ogg", ".wma"}
            Case "archives"
                extensions = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2"}
            Case "executables"
                extensions = {".exe", ".msi", ".bat", ".cmd", ".com", ".scr"}
            Case "temporary"
                extensions = {".tmp", ".temp", ".cache"}
            Case Else
                Return files ' Return empty list for special categories
        End Select

        ' Scan common user directories
        Dim searchPaths As String() = {
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
            Environment.GetFolderPath(Environment.SpecialFolder.MyMusic),
            Environment.GetFolderPath(Environment.SpecialFolder.MyVideos),
            System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")
        }

        For Each searchPath As String In searchPaths
            If Directory.Exists(searchPath) Then
                ScanDirectory(searchPath, extensions, files, category)
            End If
        Next

        ' Handle special categories
        If category.ToLower() = "large" Then
            Return files.Where(Function(f) f.Length > 100 * 1024 * 1024).ToList() ' >100MB
        ElseIf category.ToLower() = "old" Then
            Return files.Where(Function(f) f.LastWriteTime < DateTime.Now.AddYears(-1)).ToList() ' >1 year old
        End If

        Return files
    End Function

    Private Sub ScanDirectory(path As String, extensions As String(), ByRef files As List(Of FileInfo), category As String)
        Try
            Dim dir As New DirectoryInfo(path)

            ' Skip protected directories entirely
            If IsProtectedPath(dir.FullName) Then
                Return
            End If

            ' Scan files in current directory
            For Each file As FileInfo In dir.GetFiles()
                Try
                    If extensions.Length = 0 OrElse extensions.Contains(file.Extension.ToLower()) Then
                        files.Add(file)
                    End If

                    ' For special categories, add all files
                    If category.ToLower() = "large" OrElse category.ToLower() = "old" Then
                        files.Add(file)
                    End If
                Catch
                    ' Skip files that can't be accessed
                    Continue For
                End Try
            Next

            ' Recursively scan subdirectories (limit depth to avoid infinite loops)
            For Each subDir As DirectoryInfo In dir.GetDirectories()
                Try
                    ' Skip system and protected directories
                    If Not subDir.Name.StartsWith("$") AndAlso
                       Not subDir.Attributes.HasFlag(FileAttributes.System) AndAlso
                       Not subDir.Attributes.HasFlag(FileAttributes.Hidden) AndAlso
                       Not IsProtectedPath(subDir.FullName) Then
                        ScanDirectory(subDir.FullName, extensions, files, category)
                    End If
                Catch
                    ' Skip directories that can't be accessed
                    Continue For
                End Try
            Next

        Catch ex As Exception
            ' Skip directories that can't be accessed
        End Try
    End Sub

    ' Determine if a path (file or directory) is in a protected system/dev location
    Private Function IsProtectedPath(inputPath As String) As Boolean
        If String.IsNullOrWhiteSpace(inputPath) Then Return False

        Dim fullPath As String
        Try
            fullPath = System.IO.Path.GetFullPath(inputPath)
        Catch
            Return False
        End Try

        Dim directoryToCheck As String = fullPath
        If File.Exists(fullPath) Then
            Try
                directoryToCheck = System.IO.Path.GetDirectoryName(fullPath)
            Catch
                Return False
            End Try
        End If

        If String.IsNullOrEmpty(directoryToCheck) Then Return False

        Dim userProfile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Dim protectedRoots As New List(Of String) From {
            Environment.GetFolderPath(Environment.SpecialFolder.Windows),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            System.IO.Path.Combine(userProfile, "source", "repos")
        }

        ' If the directory is under any protected root, block it
        For Each root In protectedRoots
            If Not String.IsNullOrEmpty(root) AndAlso IsInsidePath(directoryToCheck, root) Then
                Return True
            End If
        Next

        ' Also block common dev/system subfolders anywhere in the path
        Dim blockedFolderNames As String() = {".git", ".vs", "bin", "obj", "node_modules", "packages"}
        If PathContainsAnyFolder(directoryToCheck, blockedFolderNames) Then
            Return True
        End If

        Return False
    End Function

    ' Check for sensitive project file types
    Private Function IsProtectedFileType(filePath As String) As Boolean
        Dim ext As String = System.IO.Path.GetExtension(filePath).ToLower()
        Dim sensitive As String() = {".sln", ".csproj", ".vbproj", ".fsproj", ".vcxproj", ".vcproj", ".proj", ".props", ".targets"}
        For Each s In sensitive
            If s = ext Then Return True
        Next
        Return False
    End Function

    Private Function IsInsidePath(childPath As String, parentPath As String) As Boolean
        Try
            Dim childFull As String = System.IO.Path.GetFullPath(childPath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)) & System.IO.Path.DirectorySeparatorChar
            Dim parentFull As String = System.IO.Path.GetFullPath(parentPath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)) & System.IO.Path.DirectorySeparatorChar
            Return childFull.StartsWith(parentFull, StringComparison.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    Private Function PathContainsAnyFolder(inputPath As String, folderNames As String()) As Boolean
        If String.IsNullOrEmpty(inputPath) Then Return False
        Dim separators() As Char = {System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar}
        Dim parts As String() = inputPath.Split(separators, StringSplitOptions.RemoveEmptyEntries)
        For Each part As String In parts
            For Each folderName As String In folderNames
                If String.Equals(part, folderName, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        ShowFilePreview()
    End Sub

    Private Sub ShowFilePreview()
        Panel1.Controls.Clear()

        If DataGridView1.SelectedRows.Count > 0 Then
            Try
                Dim filePath As String = DataGridView1.SelectedRows(0).Cells("Path").Value.ToString()
                Dim fileExt As String = Path.GetExtension(filePath).ToLower()

                ' Create info label
                Dim infoLabel As New Label()
                infoLabel.Text = $"File: {Path.GetFileName(filePath)}" & vbCrLf &
                               $"Size: {DataGridView1.SelectedRows(0).Cells("Size").Value} MB" & vbCrLf &
                               $"Modified: {DataGridView1.SelectedRows(0).Cells("Modified").Value}"
                infoLabel.Dock = DockStyle.Top
                infoLabel.Height = 60
                infoLabel.BackColor = Color.White
                Panel1.Controls.Add(infoLabel)

                ' Show preview based on file type
                If {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"}.Contains(fileExt) Then
                    ShowImagePreview(filePath)
                ElseIf {".txt", ".log", ".ini", ".cfg", ".xml", ".json", ".csv"}.Contains(fileExt) Then
                    ShowTextPreview(filePath)
                ElseIf {".pdf"}.Contains(fileExt) Then
                    ShowPdfPreview(filePath)
                ElseIf {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"}.Contains(fileExt) Then
                    ShowOfficePreview(filePath)
                ElseIf {".mp3", ".wav", ".m4a"}.Contains(fileExt) Then
                    ShowAudioPreview(filePath)
                ElseIf {".mp4", ".avi", ".mkv", ".mov"}.Contains(fileExt) Then
                    ShowVideoPreview(filePath)
                Else
                    ShowGenericPreview(filePath)
                End If

            Catch ex As Exception
                Dim errorLabel As New Label()
                errorLabel.Text = "Error loading preview: " & ex.Message
                errorLabel.Dock = DockStyle.Fill
                errorLabel.TextAlign = ContentAlignment.MiddleCenter
                errorLabel.ForeColor = Color.Red
                Panel1.Controls.Add(errorLabel)
            End Try
        End If
    End Sub

    Private Sub ShowImagePreview(filePath As String)
        Try
            ' Use a copy approach to avoid file locking
            Dim pictureBox As New PictureBox()

            ' Load image using a stream that gets disposed immediately
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim originalImage As Image = Image.FromStream(fs)
                ' Create a copy so we can dispose the original and release the file handle
                pictureBox.Image = New Bitmap(originalImage)
                originalImage.Dispose()
            End Using

            pictureBox.SizeMode = PictureBoxSizeMode.Zoom
            pictureBox.Dock = DockStyle.Fill
            pictureBox.Cursor = Cursors.Hand ' Show that it's clickable

            ' Add click event to open image externally
            AddHandler pictureBox.Click, Sub()
                                             TryOpenExternal(filePath)
                                         End Sub

            ' Add double-click event for better UX
            AddHandler pictureBox.DoubleClick, Sub()
                                                   TryOpenExternal(filePath)
                                               End Sub

            ' Add tooltip to inform user
            Dim toolTip As New ToolTip()
            toolTip.SetToolTip(pictureBox, "Click to open in default image viewer")

            Panel1.Controls.Add(pictureBox)

        Catch ex As Exception
            ' If image can't be loaded, show error message
            Dim errorLabel As New Label()
            errorLabel.Text = "Cannot preview this image file" & vbCrLf & ex.Message
            errorLabel.Dock = DockStyle.Fill
            errorLabel.TextAlign = ContentAlignment.MiddleCenter
            errorLabel.ForeColor = Color.Red
            Panel1.Controls.Add(errorLabel)
        End Try
    End Sub

    Private Sub ShowTextPreview(filePath As String)
        Try
            Dim textBox As New TextBox()
            textBox.Multiline = True
            textBox.ScrollBars = ScrollBars.Both
            textBox.ReadOnly = True
            textBox.Dock = DockStyle.Fill
            textBox.Font = New Font("Consolas", 9)

            Dim fileInfo As New FileInfo(filePath)
            If fileInfo.Length > 50000 Then ' Limit to 50KB for performance
                textBox.Text = File.ReadAllText(filePath).Substring(0, 50000) & vbCrLf & vbCrLf & "... (File truncated for preview)"
            Else
                textBox.Text = File.ReadAllText(filePath)
            End If
            Panel1.Controls.Add(textBox)
        Catch
            Dim errorLabel As New Label()
            errorLabel.Text = "Cannot preview this text file"
            errorLabel.Dock = DockStyle.Fill
            errorLabel.TextAlign = ContentAlignment.MiddleCenter
            errorLabel.ForeColor = Color.Red
            Panel1.Controls.Add(errorLabel)
        End Try
    End Sub

    Private Sub ShowPdfPreview(filePath As String)
        Try
            ' Create a panel with PDF info and open button
            Dim pdfPanel As New Panel()
            pdfPanel.Dock = DockStyle.Fill
            pdfPanel.BackColor = Color.WhiteSmoke

            ' PDF icon or placeholder
            Dim pdfIcon As New Label()
            pdfIcon.Text = "📄" ' PDF emoji
            pdfIcon.Font = New Font("Arial", 48)
            pdfIcon.TextAlign = ContentAlignment.MiddleCenter
            pdfIcon.Dock = DockStyle.Top
            pdfIcon.Height = 100
            pdfPanel.Controls.Add(pdfIcon)

            ' PDF info
            Dim pdfInfo As New Label()
            pdfInfo.Text = "PDF Document" & vbCrLf & "Click 'Open External' to view in default PDF viewer"
            pdfInfo.TextAlign = ContentAlignment.MiddleCenter
            pdfInfo.Dock = DockStyle.Fill
            pdfPanel.Controls.Add(pdfInfo)

            ' Open button
            Dim openBtn As New Button()
            openBtn.Text = "Open External"
            openBtn.Dock = DockStyle.Bottom
            openBtn.Height = 30
            AddHandler openBtn.Click, Sub() TryOpenExternal(filePath)
            pdfPanel.Controls.Add(openBtn)

            Panel1.Controls.Add(pdfPanel)

        Catch ex As Exception
            ShowGenericPreview(filePath)
        End Try
    End Sub

    Private Sub ShowOfficePreview(filePath As String)
        Try
            Dim fileExt As String = Path.GetExtension(filePath).ToLower()
            Dim officePanel As New Panel()
            officePanel.Dock = DockStyle.Fill
            officePanel.BackColor = Color.WhiteSmoke

            ' Office icon
            Dim officeIcon As New Label()
            Select Case fileExt
                Case ".doc", ".docx"
                    officeIcon.Text = "📝" ' Word document
                Case ".xls", ".xlsx"
                    officeIcon.Text = "📊" ' Excel spreadsheet
                Case ".ppt", ".pptx"
                    officeIcon.Text = "📺" ' PowerPoint presentation
                Case Else
                    officeIcon.Text = "📄"
            End Select

            officeIcon.Font = New Font("Arial", 48)
            officeIcon.TextAlign = ContentAlignment.MiddleCenter
            officeIcon.Dock = DockStyle.Top
            officeIcon.Height = 100
            officePanel.Controls.Add(officeIcon)

            ' Office info
            Dim officeInfo As New Label()
            officeInfo.Text = $"{fileExt.ToUpper()} Document" & vbCrLf & "Click 'Open External' to view in default application"
            officeInfo.TextAlign = ContentAlignment.MiddleCenter
            officeInfo.Dock = DockStyle.Fill
            officePanel.Controls.Add(officeInfo)

            ' Open button
            Dim openBtn As New Button()
            openBtn.Text = "Open External"
            openBtn.Dock = DockStyle.Bottom
            openBtn.Height = 30
            AddHandler openBtn.Click, Sub() TryOpenExternal(filePath)
            officePanel.Controls.Add(openBtn)

            Panel1.Controls.Add(officePanel)

        Catch ex As Exception
            ShowGenericPreview(filePath)
        End Try
    End Sub

    Private Sub ShowAudioPreview(filePath As String)
        Try
            Dim audioPanel As New Panel()
            audioPanel.Dock = DockStyle.Fill
            audioPanel.BackColor = Color.LightBlue

            ' Audio icon
            Dim audioIcon As New Label()
            audioIcon.Text = "🎵"
            audioIcon.Font = New Font("Arial", 48)
            audioIcon.TextAlign = ContentAlignment.MiddleCenter
            audioIcon.Dock = DockStyle.Top
            audioIcon.Height = 100
            audioPanel.Controls.Add(audioIcon)

            ' Audio info
            Dim audioInfo As New Label()
            audioInfo.Text = "Audio File" & vbCrLf & "Click 'Play' to open in default media player"
            audioInfo.TextAlign = ContentAlignment.MiddleCenter
            audioInfo.Dock = DockStyle.Fill
            audioPanel.Controls.Add(audioInfo)

            ' Play button
            Dim playBtn As New Button()
            playBtn.Text = "Play"
            playBtn.Dock = DockStyle.Bottom
            playBtn.Height = 30
            AddHandler playBtn.Click, Sub() TryOpenExternal(filePath)
            audioPanel.Controls.Add(playBtn)

            Panel1.Controls.Add(audioPanel)

        Catch ex As Exception
            ShowGenericPreview(filePath)
        End Try
    End Sub

    Private Sub ShowVideoPreview(filePath As String)
        Try
            Dim videoPanel As New Panel()
            videoPanel.Dock = DockStyle.Fill
            videoPanel.BackColor = Color.LightCoral

            ' Video icon
            Dim videoIcon As New Label()
            videoIcon.Text = "🎬"
            videoIcon.Font = New Font("Arial", 48)
            videoIcon.TextAlign = ContentAlignment.MiddleCenter
            videoIcon.Dock = DockStyle.Top
            videoIcon.Height = 100
            videoPanel.Controls.Add(videoIcon)

            ' Video info
            Dim videoInfo As New Label()
            videoInfo.Text = "Video File" & vbCrLf & "Click 'Play' to open in default media player"
            videoInfo.TextAlign = ContentAlignment.MiddleCenter
            videoInfo.Dock = DockStyle.Fill
            videoPanel.Controls.Add(videoInfo)

            ' Play button
            Dim playBtn As New Button()
            playBtn.Text = "Play"
            playBtn.Dock = DockStyle.Bottom
            playBtn.Height = 30
            AddHandler playBtn.Click, Sub() TryOpenExternal(filePath)
            videoPanel.Controls.Add(playBtn)

            Panel1.Controls.Add(videoPanel)

        Catch ex As Exception
            ShowGenericPreview(filePath)
        End Try
    End Sub

    Private Sub ShowGenericPreview(filePath As String)
        Try
            Dim genericPanel As New Panel()
            genericPanel.Dock = DockStyle.Fill
            genericPanel.BackColor = Color.LightGray

            ' Generic file icon
            Dim fileIcon As New Label()
            fileIcon.Text = "📄"
            fileIcon.Font = New Font("Arial", 48)
            fileIcon.TextAlign = ContentAlignment.MiddleCenter
            fileIcon.Dock = DockStyle.Top
            fileIcon.Height = 100
            genericPanel.Controls.Add(fileIcon)

            ' File info
            Dim fileInfo As New FileInfo(filePath)
            Dim genericInfo As New Label()
            genericInfo.Text = $"File Type: {fileInfo.Extension.ToUpper()}" & vbCrLf &
                              $"Size: {Math.Round(fileInfo.Length / (1024.0 * 1024.0), 2)} MB" & vbCrLf &
                              $"Created: {fileInfo.CreationTime:yyyy-MM-dd}" & vbCrLf &
                              $"Modified: {fileInfo.LastWriteTime:yyyy-MM-dd}" & vbCrLf & vbCrLf &
                              "Preview not available" & vbCrLf & "Click 'Open External' to view"
            genericInfo.TextAlign = ContentAlignment.MiddleCenter
            genericInfo.Dock = DockStyle.Fill
            genericPanel.Controls.Add(genericInfo)

            ' Open button
            Dim openBtn As New Button()
            openBtn.Text = "Open External"
            openBtn.Dock = DockStyle.Bottom
            openBtn.Height = 30
            AddHandler openBtn.Click, Sub() TryOpenExternal(filePath)
            genericPanel.Controls.Add(openBtn)

            Panel1.Controls.Add(genericPanel)

        Catch ex As Exception
            Dim errorLabel As New Label()
            errorLabel.Text = "Error displaying file information"
            errorLabel.Dock = DockStyle.Fill
            errorLabel.TextAlign = ContentAlignment.MiddleCenter
            errorLabel.ForeColor = Color.Red
            Panel1.Controls.Add(errorLabel)
        End Try
    End Sub

End Class