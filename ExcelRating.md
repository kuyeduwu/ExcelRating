# Excel模拟五星好评

### 如何使用

1. 选取任意连续的5列用来放置星星（假设为C-G列，A列用来写问题）。
2. 在C-G列每一列输入一个空心五角星，输入方法：按住ALT键，在小键盘区输41454，松开ALT键）。
3. 缩小C-G列列宽，使其正好等于一个星星的宽度。
3. ALT+F11打开VBA编辑器。
4. 在`Microsoft Excel Objects`下面，双击`Sheet1`。
4. 在代码编辑区输入下面的代码。
5. 关闭代码编辑器。
6. 评级的时候，只需要双击对应的星星，例如，如要打三星，只需要双击在E列的星星。

### 注意
* 在上面的例子中，五个星星为一组，其实一组里面星星的数量可以随意。
* 在Excel中，同一行只能出现一组星星，不同行之间的星星相互之间不会影响。
* 可以通过调整C-G列的字体颜色来调整星星的颜色。
* 某组星星，一旦进行了评级操作，则最低评级为一星，如果要恢复为未评级之前的状态，需要手动将星星替换。

```vb
'coded by: kuyeduwu
'coded on: 04/26/16
'website: http://www.myslsh.com
'GitHub: https://github.com/kuyeduwu/ExcelRating

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean) '由Excel的双击事件触发代码

    If Target.Text = "☆" Or Target.Text = "★" Then
        '如果当前双击的内容是五角星，则将当前双击位置以及之前的五角星全部改为实心。
        Let r = Target.Row
        Let cb = Target.Column
        Let ca = Target.Column + 1
        Do
            Cells(r, cb).Replace "☆", "★"
            cb = cb - 1
        Loop Until Cells(r, cb) <> "☆" And Cells(r, cb) <> "★"
        '将当前双击位置之后的五角星全部换成空心。
        Do
            Cells(r, ca).Replace "★", "☆"
            ca = ca + 1
        Loop Until Cells(r, ca) <> "☆" And Cells(r, ca) <> "★"
        '取消系统默认的双击效果。
        Cancel = True
    End If
End Sub

```
