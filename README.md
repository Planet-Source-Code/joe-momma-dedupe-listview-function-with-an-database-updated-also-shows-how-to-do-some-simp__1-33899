<div align="center">

## Dedupe ListView Function with an Database\. UPDATED\! Also shows how to do some simple math


</div>

### Description

My Code commenting sucked, My bad. So heres the update! This funtion is meant to be used when passing values from another control to a listview control with a database. I needed a way to keep duplicate entries away in the recieving Listview. The funtion is called after the row insertion on a double click event or Drag and drop.It works because the index is set dynamicaly by the database values.If used out of context it WILL delete everything, which in fact, does suck. I used LCase incase you want to pass a nonnumerical value(which i did and then did not), but you can remove it, no harm, no foul. Hope this helps. Feedback is very welcome. Please Feedback, Please.
 
### More Info
 
Example of use: RS is a record set in case you needed to know.

With rs ' Recordset

On Error Resume Next

Do While Not .EOF  'list products of a specific category on ListView starts loop

Set rs = ListView2.ListItems.Add(, , !Productid, 1, 1)

rs.SubItems(1) = !ProductName

rs.SubItems(2) = !Unit

rs.SubItems(3) = Format(!PRICE, "Currency")

rs.SubItems(4) = Format(!Extcost, "Currency")

rs.SubItems(5) = Format(!Retail, "Currency")

rs.SubItems(6) = Format(!Add)

cTotal = cTotal + CCur(rs.SubItems(3)) * (rs.SubItems(6)) ' here the math part , really simple CCur sets it as Currency

rs.SubItems(7) = Format(cTotal, "Currency")

.MoveNext ' Restarts loop till EOF

Loop

.Close

End With

Call RemoveDupes(ListView2)

ListView2.Refresh

End Sub

DO NOT USE OUT OF CONTEXT> IF YOU NEED A REGULAR DEDUPE FUNCTION EMAIL ME> YOU ARE MY FRIEND> I WILL HELP YOU>MABEY


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joe Momma](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joe-momma.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joe-momma-dedupe-listview-function-with-an-database-updated-also-shows-how-to-do-some-simp__1-33899/archive/master.zip)





### Source Code

```
Public Sub RemoveDupes(listview As listview)
Dim i As Integer
Dim x As Integer
For i = listview.ListItems.Count To 1 Step -1 ' Works backwards avoiding error
 For x = listview.ListItems.Count To 1 Step -1 ' same
 If i = x Then GoTo Nextx 'adds item if not in recieving listview.Once one is in there it will actually start the below routine
 If LCase(listview.ListItems(x)) = LCase(listview.ListItems(i)) Then
 listview.ListItems.Remove (i) 'that should be clear
 listview.Refresh 'Better to be safe then dead
 i = listview.ListItems.Count 'sets i to current record count so you dont get an error when the loop starts again
 End If
Nextx:
 Next x
Next i
End Sub
```

