Attribute VB_Name = "Simulator"
Public Type Node 'Data for out points
X As Double 'Position X
Y As Double 'Position Y
Xkin As Double 'Kinectic energy X
Ykin As Double 'Kinectic energy Y
mass As Single 'Its weight
Bouce As Integer 'How well it bouces
locked As Boolean 'If position is locked
End Type

Public Type Link 'Data for the links joing the points
Node1 As Integer 'The point it conects to
Node2 As Integer 'The point it conects to
Lenth As Integer 'How long the link is
flex As Integer 'How flexibe it is
Broken As Boolean 'If its broken
breakpoint As Integer 'Force at wich the link breaks
Stress As Integer 'How much force is the link stading agenst
Indestuctable As Boolean 'Link never breaks
rope As Boolean
Active As Boolean 'Show link and sumulate the link and the conected nodes
End Type
'Enviroment
Dim AirRes As Single
Dim Gravity As Single
Global SimTime As Long
Dim ModelName As String

'Someware to store the model
Dim Node(0 To 1000) As Node
Dim Link(0 To 1000) As Link

Public Sub SimulateFrame()
On Error GoTo napaka
Dim ty As Double, tx As Double, dy As Double, dx As Double
For i = 0 To 1000
    If Link(i).Active = True Then
        'Caculate the stress on the link
        ForceBetwenPoints Node(Link(i).Node1).X, Node(Link(i).Node1).Y, Node(Link(i).Node2).X, Node(Link(i).Node2).Y, Link(i).Lenth, Link(i).flex, Link(i).rope, Link(i).Stress, tx, ty, dx, dy
        'Caculate the force from the tenchion in to kinetic energy
        If Link(i).Broken = False Then ' Check if link is broken
            Node(Link(i).Node1).Xkin = Node(Link(i).Node1).Xkin + (tx / Node(Link(i).Node1).mass)
            Node(Link(i).Node2).Xkin = Node(Link(i).Node2).Xkin + (dx / Node(Link(i).Node2).mass)
            Node(Link(i).Node1).Ykin = Node(Link(i).Node1).Ykin + (ty / Node(Link(i).Node1).mass)
            Node(Link(i).Node2).Ykin = Node(Link(i).Node2).Ykin + (dy / Node(Link(i).Node2).mass)
            If (tx + ty) / 2 > Link(i).breakpoint And Link(i).Indestuctable = False Then BreakLink (i) 'Check if a destructable link is holding a big enugh force to break
        End If
        Node(Link(i).Node1).Ykin = Node(Link(i).Node1).Ykin + Gravity 'Add gravity
        Node(Link(i).Node2).Ykin = Node(Link(i).Node2).Ykin + Gravity 'Add gravity
        'Calculate air resistance
        Node(Link(i).Node1).Xkin = Node(Link(i).Node1).Xkin * AirRes
        Node(Link(i).Node1).Ykin = Node(Link(i).Node1).Ykin * AirRes
        Node(Link(i).Node2).Xkin = Node(Link(i).Node2).Xkin * AirRes
        Node(Link(i).Node2).Ykin = Node(Link(i).Node2).Ykin * AirRes
        
        
        'Calculate kinetic energy in to movement
        If Node(Link(i).Node1).locked = False Then ' Check if node is not locked
            Node(Link(i).Node1).X = Node(Link(i).Node1).X + Node(Link(i).Node1).Xkin / 2
            Node(Link(i).Node1).Y = Node(Link(i).Node1).Y + Node(Link(i).Node1).Ykin / 2
        Else
        Node(Link(i).Node1).Xkin = 0
        Node(Link(i).Node1).Ykin = 0
        End If
        If Node(Link(i).Node2).locked = False Then ' Check if node is not locked
            Node(Link(i).Node2).X = Node(Link(i).Node2).X + Node(Link(i).Node2).Xkin / 2
            Node(Link(i).Node2).Y = Node(Link(i).Node2).Y + Node(Link(i).Node2).Ykin / 2
        Else
        Node(Link(i).Node2).Xkin = 0
        Node(Link(i).Node2).Ykin = 0
        End If
    End If
Next i
SimTime = SimTime + 1
Exit Sub
napaka:
Form1.Timer1.Enabled = False
MsgBox "Simulator Error : " & Err.Number & " [ " & Err.Description & " ]", vbCritical, "ERROR !!!"
End Sub

Private Sub BreakLink(LinkNum As Integer)
Link(LinkNum).Broken = True
End Sub

Public Sub Render(img As Object)
On Error GoTo napaka
For i = 0 To 1000
    If Link(i).Active = True And Link(i).Broken = False Then
    If Form1.Cstress = 1 Then
        img.Line (Node(Link(i).Node1).X, Node(Link(i).Node1).Y)-(Node(Link(i).Node2).X, Node(Link(i).Node2).Y), StressColor(i)
    Else
        img.Line (Node(Link(i).Node1).X, Node(Link(i).Node1).Y)-(Node(Link(i).Node2).X, Node(Link(i).Node2).Y), RGB(150, 150, 150)
    End If
    
    img.DrawWidth = 5
    img.ForeColor = vbWhite
    If Form1.Cnode = 1 Then
        If Node(Link(i).Node1).locked = flase Then
            img.PSet (Node(Link(i).Node1).X, Node(Link(i).Node1).Y), vbGreen
            If Form1.Cnodeind = 1 Then img.Print Str(Link(i).Node1)
        Else
            img.PSet (Node(Link(i).Node1).X, Node(Link(i).Node1).Y), vbRed
            If Form1.Cnodeind = 1 Then img.Print Str(Link(i).Node1)
        End If
    
        If Node(Link(i).Node2).locked = False Then
            img.PSet (Node(Link(i).Node2).X, Node(Link(i).Node2).Y), vbGreen
            If Form1.Cnodeind = 1 Then img.Print Str(Link(i).Node2)
        Else
            img.PSet (Node(Link(i).Node2).X, Node(Link(i).Node2).Y), vbRed
            If Form1.Cnodeind = 1 Then img.Print Str(Link(i).Node2)
        End If
    End If
    img.ForeColor = vbGreen
    img.DrawWidth = 2
    End If
Next i
Exit Sub
napaka:
Form1.Timer1.Enabled = False
MsgBox "Render Error : " & Err.Number & " [ " & Err.Description & " ]", vbCritical, "ERROR !!!"
End Sub

Public Sub UpdateStat()
Form1.SimT = (SimTime * 0.02) & " s"
Form1.StatDisp = "Node Count : " & CountNodes & vbNewLine & "Link Count : " & CountLinks
End Sub


Public Function CountLinks() As Integer
For i = 0 To 1000
If Link(i).Active = True Then CountLinks = CountLinks + 1
Next i
End Function

Public Function CountNodes() As Integer
For i = 0 To 1000
If Node(i).X <> 0 And Node(i).Y <> 0 Then CountNodes = CountNodes + 1
Next i
End Function


Private Function StressColor(LinkNum)
If Link(LinkNum).Stress > 0 Then
StressColor = RGB(100 + Link(LinkNum).Stress, Over0(100 - Abs(Link(LinkNum).Stress)), Over0(100 - Abs(Link(LinkNum).Stress)))
Else
StressColor = RGB(Over0(100 - Abs(Link(LinkNum).Stress)), Over0(100 - Abs(Link(LinkNum).Stress)), Over0(100 + Abs(Link(LinkNum).Stress)))
End If
End Function


Public Sub SetEnviroment(AirR, Grav)
AirRes = Val(AirR)
Gravity = Val(Grav)
End Sub


Public Sub SetNode(ind, data As Node)
Node(ind) = data
End Sub

Public Sub SetLink(ind, data As Link)
Link(ind) = data
End Sub


Public Sub LoadModel(FileName As String)
Open FileName For Input As #1
Line Input #1, vrst
ModelName = vrst
Line Input #1, vrst
Gravity = vrst
Line Input #1, vrst
AirRes = vrst
Line Input #1, vrst

Do Until vrst = "# Links"
    Line Input #1, vrst
    If i > 1000 Or vrst = "# Links" Then Exit Do
    Node(i).X = CLng(vrst)
    Line Input #1, vrst
    Node(i).Y = CLng(vrst)
    Line Input #1, vrst
    Node(i).mass = Int(vrst)
    Line Input #1, vrst
    Node(i).Bouce = Int(vrst)
    Line Input #1, vrst
    Node(i).locked = TrueOrFalse(vrst)
    i = i + 1
Loop
i = 0
Do Until EOF(1)
    Line Input #1, vrst
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
