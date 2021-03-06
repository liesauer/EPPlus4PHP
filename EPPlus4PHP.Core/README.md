# EPPlus4PHP.Core

## TODO

- [x] 文档
  - [ ] 内存编辑
  - [x] 保存
  - [x] 另存为
- [x] 工作表
  - [x] 创建
  - [x] 删除
  - [x] 自动创建（不存在时，仅针对表名索引）
  - [x] 移动
- [x] 单元格定位
  - [x] 单选
  - [x] 列选
    - [x] 简易格式
  - [x] 行选
    - [x] 数字索引格式
  - [x] 窗选/框选
  - [x] 多选
  - [x] 全选
  - [x] 工作区
- [x] 数据读写
  - [x] 加入数据行
  - [x] 加入数据列
  - [x] 插入数据行
  - [x] 插入数据列
- [ ] 样式
  - [x] 字体
    - [x] 字号
    - [x] 颜色
    - [x] 加粗
    - [x] 斜体
    - [x] 下划线
    - [x] 字体外部加载
  - [ ] 填充
    - [x] 背景颜色
    - [ ] 图案样式
    - [ ] 渐变颜色
  - [x] 边框
  - [x] 单元格合并
  - [x] 对齐
    - [x] 水平对齐
    - [x] 垂直对齐
- [x] 数字格式
- [ ] 条件格式
- [ ] 数据验证
- [x] 备注
  - [ ] 富文本
- [x] 公式
  - [x] [内置公式](supported-builtin-functions.md)
  - [x] 自定义公式
  - [x] R1C1
- [ ] 图表
  - [ ] 柱状图
    - [ ] 簇状柱形图（ColumnClustered）
    - [ ] 堆积柱形图（ColumnStacked）
    - [ ] 百分比堆积柱形图（ColumnStacked100）
    - [ ] 三维簇状柱形图（ColumnClustered3D）
    - [ ] 三维堆积柱形图（ColumnStacked3D）
    - [ ] 三维百分比堆积柱形图（ColumnStacked1003D）
    - [ ] 三维柱形图（Column3D）
  - [ ] 折线图
    - [ ] 折线图（Line）
    - [ ] 堆积折线图（LineStacked）
    - [ ] 百分比堆积折线图（LineStacked100）
    - [ ] 带数据标记的折线图（LineMarkers）
    - [ ] 带标记的堆积折线图（LineMarkersStacked）
    - [ ] 带数据标记的百分比堆积折线图（LineMarkersStacked100）
    - [ ] 三维折线图（Line3D）
  - [ ] 饼图
    - [ ] 饼图（Pie）
    - [ ] 三维饼图（Pie3D）
    - [ ] 复合饼图（PieOfPie）
    - [ ] 复合条饼图（BarOfPie）
    - [ ] 圆环图（Doughnut）
  - [ ] 条形图
    - [ ] 簇状条形图（BarClustered）
    - [ ] 堆积条形图（BarStacked）
    - [ ] 百分比堆积条形图（BarStacked100）
    - [ ] 三维簇状条形图（BarClustered3D）
    - [ ] 三维堆积条形图（BarStacked3D）
    - [ ] 三维百分比堆积条形图（BarStacked1003D）
  - [ ] 面积图
    - [ ] 面积图（Area）
    - [ ] 堆积面积图（AreaStacked）
    - [ ] 百分比堆积面积图（AreaStacked100）
    - [ ] 三维面积图（）
    - [ ] 三维堆积面积图（AreaStacked3D）
    - [ ] 三维百分比堆积面积图（AreaStacked1003D）
  - [ ] XY散点图
    - [ ] 散点图（XYScatter）
    - [ ] 带平滑线和数据标记的散点图（XYScatterSmooth）
    - [ ] 带平滑线的散点图（XYScatterSmoothNoMarkers）
    - [ ] 带直线和数据标记的散点图（XYScatterLines）
    - [ ] 带直线的散点图（XYScatterLinesNoMarkers）
    - [ ] 气泡图（Bubble）
    - [ ] 三维气泡图（Bubble3DEffect）
  - [ ] 股价图
    - [ ] 盘高-盘底-收盘图（StockHLC）
    - [ ] 开盘-盘高-盘底-收盘图（StockOHLC）
    - [ ] 成交量-盘高-盘底-收盘图（StockVHLC）
    - [ ] 成交量-开盘-盘高-盘底-收盘图（StockVOHLC）
  - [ ] 曲面图
    - [ ] 三维曲面图（Surface）
    - [ ] 三维曲面图（框架图）（SurfaceWireframe）
    - [ ] 曲面图（SurfaceTopView）
    - [ ] 曲面图（俯视图框架图）（SurfaceTopViewWireframe）
  - [ ] 雷达图
    - [ ] 雷达图（Radar）
    - [ ] 带数据标记的雷达图（RadarMarkers）
    - [ ] 填充雷达图（RadarFilled）
  - [ ] 组合
    - [ ] 簇状柱形图-折线图
    - [ ] 簇状柱形图-次坐标轴上的折线图
    - [ ] 堆积面积图-簇状柱形图
    - [ ] 自定义组合
  - [ ] 更多
    - [ ] PieExploded3D
    - [ ] ConeBarClustered
    - [ ] ConeBarStacked
    - [ ] ConeBarStacked100
    - [ ] ConeCol
    - [ ] ConeColClustered
    - [ ] ConeColStacked
    - [ ] ConeColStacked100
    - [ ] CylinderBarClustered
    - [ ] CylinderBarStacked
    - [ ] CylinderBarStacked100
    - [ ] CylinderCol
    - [ ] CylinderColClustered
    - [ ] CylinderColStacked
    - [ ] CylinderColStacked100
    - [ ] DoughnutExploded
    - [ ] PieExploded
    - [ ] PyramidBarClustered
    - [ ] PyramidBarStacked
    - [ ] PyramidBarStacked100
    - [ ] PyramidCol
    - [ ] PyramidColClustered
    - [ ] PyramidColStacked
    - [ ] PyramidColStacked100
- [ ] 加密保护
  - [x] 加密
  - [ ] 保护
- [ ] VBA
- [ ] 更多
