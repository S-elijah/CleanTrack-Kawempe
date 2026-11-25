Attribute VB_Name = "CleanTrack_Completion"
Option Compare Database
Option Explicit
'------------------------------------------------------------------
' CleanTrack Kawempe - completes the missing data and relationships
' HOW TO RUN: open CleanTrack_Project Kawempe.accdb (keep a backup copy),
' press Alt+F11, File > Import File... choose this .bas, then
' press Ctrl+G and type:  CompleteCleanTrack   (Enter)
' Existing assignments are preserved and checked per report on rerun.
'------------------------------------------------------------------
Public Sub CompleteCleanTrack()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim coord As Long, teamA As Long, teamB As Long, r As Long

    ' 1. Role names must match the design (Resident, not Residents)
    db.Execute "UPDATE Users SET UserRole='Resident' WHERE UserRole='Residents'", dbFailOnError

    ' 2. The design has a CollectionTeam role but no such users exist yet
    AddUser db, "Ssemakula Ivan", "0772100200", "CollectionTeam"
    AddUser db, "Nakato Juliet", "0701300400", "CollectionTeam"

    ' 3. CollectionDay is empty for every parish
    db.Execute "UPDATE ServiceAreas SET CollectionDay='Monday'    WHERE AreaID=1 AND CollectionDay IS NULL", dbFailOnError
    db.Execute "UPDATE ServiceAreas SET CollectionDay='Tuesday'   WHERE AreaID=2 AND CollectionDay IS NULL", dbFailOnError
    db.Execute "UPDATE ServiceAreas SET CollectionDay='Wednesday' WHERE AreaID=3 AND CollectionDay IS NULL", dbFailOnError
    db.Execute "UPDATE ServiceAreas SET CollectionDay='Thursday'  WHERE AreaID=4 AND CollectionDay IS NULL", dbFailOnError
    db.Execute "UPDATE ServiceAreas SET CollectionDay='Friday'    WHERE AreaID=5 AND CollectionDay IS NULL", dbFailOnError

    ' 4. Add a sample assignment only when that report has none.
    ' This allows a rerun to finish after a partial failure.
    coord = Nz(DLookup("UserID", "Users", "UserRole='Coordinator'"), 0)
    teamA = Nz(DLookup("UserID", "Users", "FullName='Ssemakula Ivan' AND UserRole='CollectionTeam'"), 0)
    teamB = Nz(DLookup("UserID", "Users", "FullName='Nakato Juliet' AND UserRole='CollectionTeam'"), 0)
    If coord > 0 And teamA > 0 And teamB > 0 Then
        For r = 1 To 3
            If DCount("*", "Reports", "ReportID=" & r) > 0 Then
                If DCount("*", "Assignments", "ReportID=" & r) = 0 Then
                    db.Execute "INSERT INTO Assignments (ReportID, CoordinatorID, TeamMemberID, AssignedAt, TargetDate) " & _
                        "VALUES (" & r & "," & coord & "," & IIf(r Mod 2 = 1, teamA, teamB) & ",Now(),Date()+3)", dbFailOnError
                End If
            End If
        Next r
    Else
        MsgBox "Sample assignments were skipped: a coordinator or collection team user was not found. Check user roles and rerun.", vbExclamation
    End If

    ' 5. Relationships still missing (every Users foreign key must be enforced)
    AddRel db, "UsersTeamAssignments", "Users", "Assignments", "UserID", "TeamMemberID"
    AddRel db, "UsersStatusEvents", "Users", "StatusEvents", "UserID", "UpdatedBy"
    AddRel db, "UsersNotifications", "Users", "Notifications", "UserID", "RecipientID"

    MsgBox "CleanTrack completion finished. Check Database Tools > Relationships.", vbInformation
End Sub

Private Sub AddUser(db As DAO.Database, nm As String, ph As String, role As String)
    If DCount("*", "Users", "FullName='" & nm & "'") = 0 Then
        db.Execute "INSERT INTO Users (FullName, PhoneNumber, UserRole, AccountStatus) " & _
                   "VALUES ('" & nm & "','" & ph & "','" & role & "','Active')", dbFailOnError
    End If
End Sub

Private Sub AddRel(db As DAO.Database, relName As String, pk As String, fk As String, pkField As String, fkField As String)
    Dim rel As DAO.Relation, f As DAO.Field, x As DAO.Relation
    For Each x In db.Relations
        If x.Name = relName Then Exit Sub
        If x.Table = pk And x.ForeignTable = fk Then
            If x.Fields.Count > 0 Then If x.Fields(0).ForeignName = fkField Then Exit Sub
        End If
    Next x
    Set rel = db.CreateRelation(relName, pk, fk, dbRelationUpdateCascade)
    Set f = rel.CreateField(pkField)
    f.ForeignName = fkField
    rel.Fields.Append f
    db.Relations.Append rel
End Sub
