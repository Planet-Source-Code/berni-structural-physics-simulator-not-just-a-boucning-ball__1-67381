Attribute VB_Name = "ModelEditor"
Const ColorLink = &HFFFFFF 'Link Color
Const ColorLinkH = &HFFFF80 'Link Handle Color
Const ColorNode = &HFF00& 'Node color
Const ColorNodeL = &HFF& 'Locked Node color
Const ColorText = &HC0C0C0 'Node index color
Const ColorLinkDraw = &HE0E0E0
Const DWLink = 2 'Link draw width
Const DWNode = 5 'Node draw width
Const DWLinkH = 10 'Link Handle draw width
Const DWLinkDraw = 5 'Drawing line for likes draw width


Dim Node(0 To 1000) As Node
Dim Link(0 To 1000) As Link

Dim MovTrue As Boolean
Dim MovIndex As Integer
Dim ReLenLinks As Boolean
Dim SelNode As Integer
Dim LinkDrw As Boolean
Dim LinkHDrw As Boolean
Dim CurX, CurY
 
Public Sub LinkHandleDrawEnable(state As Boolean)
LinkHDrw = state
End Sub

Public Sub ReLenLinkEnable(state As Boolean)
ReLenLinks = state
End Sub
 
Public Sub LinkDrawEnable(state As Boolean)
LinkDrw = state
End Sub
 
Public Function GetLinkDat(ind) As Link
GetLinkDat = Link(ind)
End Function

Public Function GetNodeDat(ind) As Node
GetNodeDat = Node(ind)
End Function

Public Sub SetLinkDat(ind, flex, breakpoint, nobreak, rope)
Link(ind).flex = flex
Link(ind).breakpoint = breakpoint
Link(ind).Indestuctable = nobreak
Link(ind).rope = rope
End Sub

Public Sub SetNodeDat(ind, mass, bounce, locked)
Node(ind).mass = mass
Node(ind).Bouce = bounce
Node(ind).locked = locked
End Sub
 
Public Function GetLinkHandle(X, Y) As Integer
GetLinkHandle = 32000
For i = 0 To 1000
lX = Node(Link(i).Node1).X - ((Node(Link(i).Node1).X - Node(Link(i).Node2).X) / 2)
lY = Node(Link(i).Node1).Y - ((Node(Link(i).Node1).Y - Node(Link(i).Node2).Y) / 2)
If Abs(lX - X) < 100 And Abs(lY - Y) < 100 And lX > 0 And lY > 0 Then GetLinkHandle = i
Next i
End Function


Public Sub StartMov(ind)
MovTrue = True
MovIndex = ind
End Sub

Public Sub EndMov()
MovTrue = False
If ReLenLinks = True Then
    For a = 0 To 1000
        If Link(a).Node1 = MovIndex Or Link(a).Node2 = MovIndex Then Link(a).Lenth = Distance(Node(Link(a).Node1).X, Node(Link(a).Node1).Y, Node(Link(a).Node2).X, Node(Link(a).Node2).Y)
    Next a
End If
End Sub

Public Sub UpdateCursor(X, Y)
CurX = X
CurY = Y
If MovTrue = True Then
Node(MovIndex).X = X
Node(MovIndex).Y = Y
End If

End Sub

Public Sub DeleteElement(X, Y)
For i = 0 To 1000
lX = Node(Link(i).Node1).X - ((Node(Link(i).Node1).X - Node(Link(i).Node2).X) / 2)
lY = Node(Link(i).Node1).Y - ((Node(Link(i).Node1).Y - Node(Link(i).Node2).Y) / 2)
If Abs(lX - X) < 100 And Abs(lY - Y) < 100 And lX > 0 And lY > 0 Then Link(i).Active = False
Next i
For i = 0 To 1000
    If Abs(Node(i).X - X) < 100 And Abs(Node(i).Y - Y) < 100 And Node(i).X > 0 And Node(i).Y > 0 Then
        Node(i).X = 0
        Node(i).Y = 0
        For a = 0 To 1000
            If Link(a).Node1 = i Or Link(a).Node2 = i Then Link(a).Active = False
        Next a
        GoTo konec
    End If
Next i

konec:
End Sub


Public Sub TrasferModelToSim()
For i = 0 To 1000
SetNode i, Node(i)
SetLink i, Link(i)
Next i
End Sub

Public Function SelectNode(X, Y) As Integer
SelectNode = 32000
SelNode = 32000
For i = 0 To 1000
If Abs(Node(i).X - X) < 100 And Abs(Node(i).Y - Y) < 100 And Node(i).X > 0 And Node(i).Y > 0 Then
SelectNode = i
SelNode = i
Exit For
End If
Next i
End Function

Public Sub AddLink(Node1, Node2, Flexing, breakpoint, Indestructable, rope)
i = FreeLinkVar
Link(i).Node1 = Node1
Link(i).Node2 = Node2
Link(i).flex = Flexing
Link(i).breakpoint = breakpoint
If Indestructable = 1 Then
Link(i).Indestuctable = True
Else
Link(i).Indestuctable = False
End If

If rope = 1 Then
Link(i).rope = True
Else
Link(i).rope = False
End If

Link(i).Active = True

Link(i).Lenth = Distance(Node(Node1).X, Node(Node1).Y, Node(Node2).X, Node(Node2).Y)

End Sub


Public Sub AddNode(X, Y, mass, Bouce, locked)
i = FreeNodeVar
Node(i).X = X
Node(i).Y = Y
Node(i).mass = mass
Node(i).Bouce = Bouce
If locked = 1 Then
Node(i).locked = True
Else
Node(i).locked = False
End If
End Sub
Public Function MaxNodeID() As Integer
For i = 0 To 1000
If Node(i).X <> 0 And Node(i).Y <> 0 Then MaxNodeID = i
Next i
End Function

Public Function MaxLinkID() As Integer
For i = 0 To 1000
If Link(i).Active = True Then MaxLinkID = i
Next i
End Function


Public Function FreeNodeVar()
i = 0
Do Until i = 1000
If Node(i).X = 0 And Node(i).Y = 0 Then Exit Do
i = i + 1
Loop
FreeNodeVar = i
End Function

Public Function FreeLinkVar()
i = 0
Do Until i = 1000
If Link(i).Active = False Then Exit Do
i = i + 1
Loop
FreeLinkVar = i
End Function
Public Sub SaveModelEdit(FileName As String)
Dim temp As String
temp = temp & ModelName & vbNewLine
temp = temp & Form2.Tgrav & vbNewLine
temp = temp & Form2.Tair & vbNewLine
temp = temp & "# Nodes" & vbNewLine
For i = 0 To MaxNodeID
    temp = temp & Node(i).X & vbNewLine
    temp = temp & Node(i).Y & vbNewLine
    temp = temp & Node(i).mass & vbNewLine
    temp = temp & Node(i).Bouce & vbNewLine
    If Node(i).locked = True Then
        temp = temp & "1" & vbNewLine
    Else
        temp = temp & "0" & vbNewLine
    End If
Next i
temp = temp & "# Links" & vbNewLine
For i = 0 To MaxLinkID
    temp = temp & Link(i).Node1 & vbNewLine
    temp = temp & Link(i).Node2 & vbNewLine
    temp = temp & Link(i).Lenth & vbNewLine
    temp = temp & Link(i).flex & vbNewLine
    temp = temp & Link(i).breakpoint & vbNewLine
    If Link(i).Indestuctable = True Then
        temp = temp & "1" & vbNewLine
    Else
        temp = temp & "0" & vbNewLine
    End If
    If Link(i).rope = True Then
        temp = temp & "1" & vbNewLine
    Else
        temp = temp & "0" & vbNewLine
    End If
Next i

Open FileName For Output As #2
Print #2, temp
Close #2
End Sub


Public Sub LoadModelEdit(FileName As String)
Open FileName For Input As #1
Line Input #1, vrst
ModelName = vrst
Line Input #1, vrst
Form2.Tgrav = vrst
Line Input #1, vrst
Form2.Tair = vrst
Line Input #1, vrst

Do Until vrst = "# Links"
    Line Input #1, vrst
    If i > 1000 Or vrst = "# Links" Then Exit Do
    Node(i).X = CLng(vrst)
    Line Input #1, vrst
    Node(i).Y = CLng(vrst)
    Line Input #1, vrst
    Node(i).mass = vrst
    Line Input #1, vrst
    Node(i).Bouce = vrst
    Line Input #1, vrst
    Node(i).locked = TrueOrFalse(vrst)
    i = i + 1
Loop
i = 0
Do Until EOF(1)
    Line Input #1, vrst
    If vrst = "" Then Exit Do
    Link(i).Node1 = Int(vrst)
    Line Input #1, vrst
    Link(i).Node2 = Int(vrst)
    Line Input #1, vrst
    Link(i).Lenth = Int(vrst)
    Line Input #1, vrst
    Link(i).flex = Int(vrst)
    Line Input #1, vrst
    Link(i).breakpoint = Int(vrst)
    Line Input #1, vrst
    Link(i).Indestuctable = TrueOrFalse(vrst)
    Line Input #1, vrst
    Link(i).rope = TrueOrFalse(vrst)
    Link(i).Active = True
    i = i + 1
    If i > 1000 Then Exit Do
Loop
Close #1
End Sub

Public Sub ClearModel()
For i = 0 To 1000
Node(i).X = 0
Node(i).Y = 0
Link(i).Active = False
Next i
End Sub

Public Sub RenderEdit(img As Object)
'On Error Resume Next
img.Cls
For i = 0 To 1000
    img.DrawWidth = DWLink
    If Link(i).Active = True And Link(i).Broken = False Then
        img.Line (Node(Link(i).Node1).X, Node(Link(i).Node1).Y)-(Node(Link(i).Node2).X, Node(Link(i).Node2).Y), ColorLink
        img.DrawWidth = DWLinkH
        If LinkHDrw = True Then img.PSet (Node(Link(i).Node1).X - ((Node(Link(i).Node1).X - Node(Link(i).Node2).X) / 2), Node(Link(i).Node1).Y - ((Node(Link(i).Node1).Y - Node(Link(i).Node2).Y)) / 2), ColorLinkH
    End If
    img.DrawWidth = DWNode
    img.ForeColor = ColorText
    If Node(i).X > 0 And Node(i).Y > 0 Then
        If Node(i).locked = flase Then
            img.PSet (Node(i).X, Node(i).Y), ColorNode
            img.Print Str(i)
        Else
            img.PSet (Node(i).X, Node(i).Y), ColorNodeL
            img.Print Str(i)
        End If
    End If
Next i
img.ForeColor = ColorLinkDraw
img.DrawWidth = DWLinkDraw
If SelNode < 1000 And LinkDrw = True Then
    img.PSet (Node(SelNode).X, Node(SelNode).Y)
    img.Line (Node(SelNode).X, Node(SelNode).Y)-(CurX, CurY)
End If
End Sub
