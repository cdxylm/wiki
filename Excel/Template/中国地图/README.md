## Excel快速制作填色地图
### 基本原理
    将矢量图片导入excel中，取消组合后得到各地区图形对象，分别填充颜色。  
    
    VBA代码：
    Sub 上色()
    Dim i, j, name
    j = 4
    
    For i = 2 To j
        name = Range("a" & i).Value
        Worksheets("地图").Shapes(name).Fill.ForeColor.RGB = Range("b" & i).DisplayFormat.Interior.Color    '图形填充颜色'
        Worksheets("地图").Shapes(name).Line.ForeColor.RGB = RGB(105, 105, 105)     '图形边框颜色'
        Worksheets("地图").Shapes(name).TextFrame2.TextRange.Characters.Text = Range("a" & i).Value & Chr(10) & Range("b" & i).Value    '图形内文字'
    Next
    MsgBox ("Ok!")
    End Sub



### 功能
    - [x] 基本功能
