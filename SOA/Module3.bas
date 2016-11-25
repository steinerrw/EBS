Attribute VB_Name = "Control"
Option Explicit
Sub Pause(sngSeconds As Single)
    Dim sngStart As Single
    
    sngStart = Timer
    Do While Timer < sngStart + sngSeconds
        DoEvents
    Loop
End Sub

Public Sub lineStat(x0, x1, y0, y1, bs, os, xs, ys)
Dim dL As Double
Dim s As Double
   dX = x1 - x0
   dY = y1 - y0
   dL = Dist(x1, y1, x0, y0)
   s = os - bs
   xs = x0 + (s * (dX / dL))
   ys = y0 + (s * (dY / dL))
End Sub

Public Sub xyzCoord(x, y, z, x0, y0, x1, y1, offset, elev, elem, dir, xx, yy, zz)
Dim dL As Double
Dim IB As Double
   dX = x - x0
   dY = y - y0
   dL = Dist(x, y, x0, y0)
   If dL = 0 Then
      dX = x - x1
      dY = y - y1
      dL = Dist(x1, y1, x, y)
   End If
   If elem = "L" Then
      xx = x + (offset * (dY / dL))
      yy = y - (offset * (dX / dL))
   Else
      IB = -1
      If dir = "L" Then
         IB = 1
      End If
      xx = x + (IB * offset * (dX / dL))
      yy = y + (IB * offset * (dY / dL))
   End If
   If offset = 0 Then
      zz = z
   Else
      zz = z + elev
   End If
End Sub


Public Sub writeRtn(xx, yy, zz, sta, offset, elev, icnt)
Dim gp3line As String
   xx = Format(xx, "######0.000")
   yy = Format(yy, "######0.000")
   zz = Format(zz, "######0.000")
   sta = Format(sta, "######0.000")
   offset = Format(offset, "####0.000")
   elev = Format(elev, "###0.000")
   Write #6, xx, yy, zz, sta, offset, elev
   Write #7, xx, yy, zz
   Write #10, sta, offset, zz
   xx = Format(xx, "######0.00")
   yy = Format(yy, "######0.00")
   zz = Format(zz, "######0.00")
   sta = Format(sta, "######0.00")
   gp3line = "   S" & "      " & icnt & "       " & yy & "       " & xx & "       " & zz & "  " & "$ X"
'   Print #20, gp3line
   Print #20, gp3line, sta, offset, elev
End Sub

Public Sub curveStat(x0, y0, x1, y1, bsta, sta, dir, rad, xc, yc)
Dim deltas As Double
Dim theta1 As Double
Dim theta2 As Double
Dim theta  As Double
Dim xt As Double
Dim yt As Double
   deltas = sta - bsta
   Call GetAng(x0, y0, x1, y1, Ang)
'   theta1 = Alignment.Rad2Deg(Ang)
   theta1 = ((Ang / 180) * PI)
   theta2 = deltas / rad
   If dir = "R" Then
      theta = theta1 - theta2
   Else
      theta = theta1 + theta2
   End If
   xt = Cos(theta)
   yt = Sin(theta)
   xc = x0 + rad * xt
   yc = y0 + rad * yt
End Sub
