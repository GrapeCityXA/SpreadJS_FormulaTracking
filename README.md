# SpreadJS_FormulaTracking
公式追踪
### SpreadJS 示例，公式追踪
该示例包括使用 SpreadJS API 的演示脚本，可用于实现公式追踪
有关 SpreadJS API 的更多信息，请参阅[SpreadJS API指南]( https://demo.grapecity.com.cn/spreadjs/help/api/) 和[帮助手册]( https://help.grapecity.com.cn/pages/viewpage.action?pageId=5963808)。


### 运行步骤
1、在开始之前，请确保您已满足以下先决条件：
要运行 SpreadJS，浏览器必须支持 HTML5，客户端导入和导出 Excel 需要 IE10及以上。
请先了解 [SpreadJS 的产品使用环境]( https://www.grapecity.com.cn/developer/spreadjs/selection-guide/product-use-environment)，并申请临时部署授权激活
安装并更新NodeJS和NPM
2、克隆或下载此代码库
3、初始化控件，并运行示例脚本
#### 控件初始化
首先，创建一个新页面，并在页面上输入以下代码：
```
<!DOCTYPE html>
    <html>
    <head>
        <title>SpreadJS HTML Test Page</title>
```
2、在页面中添加对 SpreadJS 的引用。代码如下。需要注意的是，SpreadJS 提供压缩过
```
//（minified）的 JavaScript 文件和和用于调试的文件：
<script src="[Your_Scripts_Path]/gc.spread.sheets.all.xxxx.min.js" type="text/javascript"></script>
```
3、添加 CSS 文件以改变Spread.JS 的外观。默认的CSS文件名为： 
gc.spread.sheets.xxxx.css，里面包含了所有的默认样式。该 CSS 文件将会影响滚动条，筛选框及其子元素，单元格和下方标签栏的样式。引入 CSS 的代码如下：
```
//<link href="[Your_CSS_Path]/gc.spread.sheets.xxxx.css" rel="stylesheet" type="text/css"/>
```
4、添加产品授权，代码为（本地测试可以不添加）：
```
GC.Spread.Sheets.LicenseKey = "xxx";
```
5. 添加控件初始化代码。本例会在一个 id 为 “ss” 的 DOM 元素上初始化 SpreadJS：
```
<script type="text/javascript">
// Add your license
// If run this in local for testing, remove or comment below code
 GC.Spread.Sheets.LicenseKey = "xxx";

// Add your code
 window.onload = function(){
var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"),{sheetCount:3});
var sheet = spread.getActiveSheet();
 }
</script>
</head>
<body>
```
6、 创建一个 id 为 “ss” 的元素，SpreadJS 将在该 DOM 中初始化：
```
<div id="ss" style="height: 500px; width: 800px"></div>
</body>
</html>
```
#### 示例代码
```
HTML：
<div class="sample-tutorial">
    <div id="ss" class="sample-spreadsheets"></div>

    <div class="options-container">
        <div class="option-row">
            <div class="inputContainer">
                <input type="file" id="fileDemo" class="input">
                <input type="button" id="loadExcel" value="import" class="button">
            </div>
            <div class="inputContainer">
                <input id="exportFileName" value="export.xlsx" class="input">
                <input type="button" id="saveExcel" value="export" class="button">
            </div>
        </div>
        <div class="option-row">
            <div class="group">
                <label>Password:
                    <input type="password" id="password">
                </label>
            </div>
        </div>
        <div class="option-row">
            <div class="inputContainer">
                <label>选择单元格获取公式依赖信息，双击依赖树结点跳转</label>
                <div>
                    <input type="button" id="trackPrecedentsCell" value="追踪引用单元格" class="button">
                </div>
                <div>
                    <input type="button" id="trackDependentsCell" value="追踪从属单元格" class="button">
                </div>
                <div>

                    <input type="button" id="trackAllCell" value="追踪所有单元格" class="button">
                </div>
            </div>
        </div>
    </div>
    <div id="ss1" class="sample-spreadsheets"></div>
</div>

CSS：
.sample-tutorial {
    position: relative;
    height: 100%;
    overflow: hidden;
}

.sample-spreadsheets {
    width: calc(100% - 280px);
    height: 50%;
    overflow: hidden;
    float: left;
}

.options-container {
    float: right;
    width: 280px;
    padding: 12px;
    height: 100%;
    box-sizing: border-box;
    background: #fbfbfb;
    overflow: auto;
}

.sample-options {
    z-index: 1000;
}

.inputContainer {
    width: 100%;
    height: auto;
    border: 1px solid #eee;
    padding: 6px 12px;
    margin-bottom: 10px;
    box-sizing: border-box;
}

.input {
    font-size: 14px;
    height: 20px;
    border: 0;
    outline: none;
    background: transparent;
}

.button {
    height: 30px;
    padding: 6px 12px;
    width: 120px;
    margin-top: 6px;
}

.group {
    padding: 12px;
}

.group input {
    padding: 4px 12px;
}

body {
    position: absolute;
    top: 0;
    bottom: 0;
    left: 0;
    right: 0;
}

JavaScript：
// Title:公式追踪
// Description：公式追踪
// Tag:公式追踪
var spreadForShow = null; //用于展示的spread对象
var maxDeep = 5; //最大深度
var firstShapeX = 20; //第一个shape开始位置
var firstShapeY = 20; //第一个shape开始位置
var rectWidth = 260; //单元格展示信息矩形宽度
var rectHeight = 60; //矩形高度
var spacingWidth = 300; //形状带空白宽度，rectWidth+横向间距
var shapeGap = 20; //shape纵向间距

var trackCellInfo, trackType, sourceSpread;


function workbookInitialized(spread) {
    if (spread) {
        spread.options.scrollByPixel = true; //像素滚动
        let sheet = spread.getActiveSheet();
        sheet.defaults.rowHeight = this.rectHeight + this.shapeGap; //方便插入行计算
        sheet.defaults.colWidth = this.spacingWidth; //方便插入列计算
        sheet.options.gridline.showVerticalGridline = false; //隐藏网格线
        sheet.options.gridline.showHorizontalGridline = false; //隐藏网格线
        // sheet.options.rowHeaderVisible = false;
        // sheet.options.colHeaderVisible = false;
        sheet.options.protectionOptions.allowEditObjects = true; //允许保护后编辑shape
        spread.getHost().addEventListener("dblclick", this.workbookDblClicked);
        sheet.options.isProtected = true;
    }
    console.log("tracker workbookInitialized " + this.trackCellInfo);
}


// 处理Shape双击事件，跳转对应单元格
function workbookDblClicked(e) {
    console.log(e)
    let self = window;
    if (!self.spreadForShow || !self.sourceSpread) {
        return;
    }
    var sheet = self.spreadForShow.getActiveSheet();
    if (!sheet) {
        return false;
    }
    let host = self.spreadForShow.getHost();
    var offset = $(host).offset(),
        left = offset.left,
        top = offset.top;
    var x = e.pageX - left,
        y = e.pageY - top;
    var hitTest = sheet.hitTest(x, y);
    if (!hitTest || !hitTest.shapeHitInfo) {
        return;
    }
    // 获取双击选中shape
    var shapes = sheet.shapes.all(),
        activeShape = null;
    for (var i = 0; i < shapes.length; i++) {
        var shape = shapes[i];
        if (shape.isSelected()) {
            activeShape = shape;
            break;
        }
    }
    if (activeShape && activeShape.type() === GC.Spread.Sheets.Shapes.AutoShapeType.rectangle) {
        let item = self.getCellInfo(activeShape.name());
        console.log(activeShape.name())
        self.sourceSpread.suspendPaint()
        let sheet = self.sourceSpread.getSheetFromName(item.sheetName);
        if (sheet) {
            self.sourceSpread.setActiveSheet(item.sheetName);
            self.sourceSpread.startSheetIndex(self.sourceSpread.getSheetIndex(item.sheetName));
            sheet.setActiveCell(item.row, item.col);
            sheet.showCell(item.row, item.col, GC.Spread.Sheets.VerticalPosition.center, GC.Spread.Sheets.HorizontalPosition.center);
        }
        self.sourceSpread.resumePaint()
    }
}

function trackCellInfoChanged(trackCellInfo, sourceSpread, spreadForShow, trackType) {
    this.trackType = trackType;
    this.trackCellInfo = trackCellInfo;
    this.sourceSpread = sourceSpread;
    this.spreadForShow = spreadForShow;
    if (trackCellInfo && sourceSpread && spreadForShow) {
        spreadForShow.suspendPaint();
        buildNodeTreeAndPaint(sourceSpread, spreadForShow, trackCellInfo);
        spreadForShow.resumePaint();
    }
}

//解析Cell信息，“SheetName*row*col”形式
function getCellInfo(cellInfo) {
    let info = cellInfo.split("*");
    return {
        sheetName: info[0],
        row: parseInt(info[1]),
        col: parseInt(info[2])
    }
}



// 递归构建追踪树
function buildNodeTreeAndPaint(spreadSource, spreadForShow, trackCellInfo) {
    let info = this.getCellInfo(trackCellInfo);
    spreadForShow.suspendPaint();
    var sheetSource = spreadSource.getSheetFromName(info.sheetName);
    var sheetForShow = spreadForShow.getActiveSheet();
    sheetForShow.shapes.clear();
    // 创建跟节点
    let rootNode = this.creatNode(info.row, info.col, sheetSource, 0, "")
    // shapeName记录单元格信息
    let name = rootNode.sheetName + "*" + rootNode.row + "*" + rootNode.col + "*" + Math.random().toString();
    // 绘制第一个根shape
    let fatherShape = this.getRectShape(sheetForShow, name, this.firstShapeX, this.firstShapeY, this.rectWidth, this.rectHeight, rootNode);

    // 双向递归追踪单元格并绘制
    if (this.trackType === "Precedents" || this.trackType === "Both") {
        this.getNodeChild(rootNode, sheetSource, "Precedents")
        console.log(rootNode)
        var deepInfo = [1];
        if (rootNode.childNodes && rootNode.childNodes.length) {
            this.paintDataTreeFromRoot(sheetForShow, rootNode, rootNode.childNodes.length, fatherShape, deepInfo);
        }
    }
    if (this.trackType === "Dependents" || this.trackType === "Both") {
        this.getNodeChild(rootNode, sheetSource, "Dependents")
        console.log(rootNode)
        var deepInfo = [1];
        if (rootNode.childNodes && rootNode.childNodes.length) {
            this.paintDataTreeFromRoot(sheetForShow, rootNode, rootNode.childNodes.length, fatherShape, deepInfo);
        }
    }

    // 显示fatherShape
    spreadForShow.options.scrollByPixel = false;
    let row = fatherShape.startRow(),
        col = fatherShape.startColumn();
    sheetForShow.setActiveCell(row, col);
    sheetForShow.showCell(row, col, GC.Spread.Sheets.VerticalPosition.top, GC.Spread.Sheets.HorizontalPosition.center);

    spreadForShow.options.scrollByPixel = true;
    spreadForShow.resumePaint();
}

// 创建节点
function creatNode(row, col, sheet, deep, trackType) {
    var node = {
        value: sheet.getValue(row, col),
        position: sheet.name() + "!" + GC.Spread.Sheets.CalcEngine.rangeToFormula(new GC.Spread.Sheets.Range(row, col, 1, 1)),
        deep: deep,
        sheetName: sheet.name(),
        row: row,
        col: col,
        trackType: trackType
    };
    return node;
}
// 递归获取子节点
function getNodeChild(rootNode, sheet, trackType) {
    let childNodeArray = [];
    let childNodes = [];
    let row = rootNode.row,
        col = rootNode.col,
        deep = rootNode.deep;
    if (trackType == "Precedents") {
        childNodes = sheet.getPrecedents(row, col);
    } else {
        childNodes = sheet.getDependents(row, col);
    }
    let self = this;
    if (childNodes.length >= 1) {
        childNodes.forEach(function(node) {
            let row = node.row,
                col = node.col,
                rowCount = node.rowCount,
                colCount = node.colCount,
                _sheet = sheet.parent.getSheetFromName(node.sheetName);
            if (rowCount > 1 || colCount > 1) {
                for (let r = row; r < row + rowCount; r++) {
                    for (let c = col; c < col + colCount; c++) {
                        let newNode = self.creatNode(r, c, sheet, deep + 1, trackType)
                        if (deep < self.maxDeep) {
                            self.getNodeChild(newNode, sheet, trackType);
                        }
                        childNodeArray.push(newNode);
                    }
                }
            } else {
                let newNode = self.creatNode(row, col, sheet, deep + 1, trackType)
                if (deep < self.maxDeep) {
                    self.getNodeChild(newNode, sheet, trackType);
                }
                childNodeArray.push(newNode);
            }
        });
    }
    rootNode.childNodes = childNodeArray;
}
// 绘制矩形shape
function getRectShape(sheetForShow, name, x, y, width, height, nodeTree) {
    var rectShape = sheetForShow.shapes.add(name, GC.Spread.Sheets.Shapes.AutoShapeType.rectangle, x, y, width, height);
    var oldStyle = rectShape.style();

    oldStyle.fill.color = "#2894FF";
    oldStyle.textEffect.font = "bold 15px Calibri";
    oldStyle.textFrame.vAlign = GC.Spread.Sheets.VerticalAlign.top;
    oldStyle.textFrame.hAlign = GC.Spread.Sheets.HorizontalAlign.left;
    if (nodeTree.deep === 0) {
        oldStyle.textEffect.color = "yellow";
    } else {
        oldStyle.textEffect.color = "white";
    }
    rectShape.style(oldStyle);
    rectShape.dynamicMove(true);

    var _description = "Value: " + nodeTree.value + "    deep:" + nodeTree.deep + "\nCell: " + nodeTree.position;
    rectShape.text(_description);

    return rectShape;
}
// 添加链接符
function getConnectorShape(sheetForShow) {
    return sheetForShow.shapes.addConnector("", GC.Spread.Sheets.Shapes.ConnectorType.elbow);
}
// 递归绘制shape
function paintDataTreeFromRoot(sheetForShow, rootNode, childLength, fatherShape, deepInfo) {
    if (!fatherShape) {
        return;
    }
    var childNodes = rootNode.childNodes;
    if (childNodes) {
        let self = this;
        var rectWidth = self.rectWidth,
            rectHeight = self.rectHeight;
        var spacingWidth = self.spacingWidth,
            shapeGap = self.shapeGap;

        for (let index = 0; index < childNodes.length; index++) {
            let nodeTree = childNodes[index];

            // 绘制shape
            var startIndex = deepInfo[nodeTree.deep] ? deepInfo[nodeTree.deep] : 0;
            var x = fatherShape.x() + spacingWidth;
            if (nodeTree.trackType == "Precedents") {
                x = fatherShape.x() - spacingWidth;
            }
            if (x < 0) {
                sheetForShow.addColumns(0, 1);
                x += sheetForShow.defaults.colWidth;
            }
            var y = self.firstShapeY + startIndex * (rectHeight + shapeGap);
            if (y < fatherShape.y()) {
                y = fatherShape.y();
                deepInfo[nodeTree.deep] = deepInfo[nodeTree.deep - 1] - 1;
            }
            if (index === 0 && y > fatherShape.y()) {
                deepInfo[nodeTree.deep - 1] = deepInfo[nodeTree.deep] + 1;
                fatherShape.y(y);
            }
            if (deepInfo[nodeTree.deep]) {
                deepInfo[nodeTree.deep]++;
            } else {
                deepInfo[nodeTree.deep] = 1;
            }
            var name = nodeTree.sheetName + "*" + nodeTree.row + "*" + nodeTree.col + "*" + Math.random().toString();
            let rectShape = self.getRectShape(sheetForShow, name, x, y, rectWidth, rectHeight, nodeTree);

            rectShape.text(rectShape.text() + "deepinfo:" + deepInfo[nodeTree.deep]);

            //绘制链接符
            var connectorShape = self.getConnectorShape(sheetForShow);
            let connectorStyle = connectorShape.style();
            if (nodeTree.trackType == "Precedents") {
                connectorStyle.line.beginArrowheadStyle = GC.Spread.Sheets.Shapes.ArrowheadStyle.triangle;
                connectorShape.startConnector({
                    name: fatherShape.name(),
                    index: 1
                });
                connectorShape.endConnector({
                    name: rectShape.name(),
                    index: 3
                });
            } else {
                connectorStyle.line.endArrowheadStyle = GC.Spread.Sheets.Shapes.ArrowheadStyle.triangle;
                connectorShape.startConnector({
                    name: fatherShape.name(),
                    index: 3
                });
                connectorShape.endConnector({
                    name: rectShape.name(),
                    index: 1
                });
            }
            connectorShape.style(connectorStyle);

            //递归绘制
            if (nodeTree.childNodes && nodeTree.childNodes.length) {
                this.paintDataTreeFromRoot(sheetForShow, nodeTree, nodeTree.childNodes.length, rectShape, deepInfo);
            }
        }
    }

}
var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), {
    calcOnDemand: true
});
var trackSpread = new GC.Spread.Sheets.Workbook(document.getElementById("ss1"));
workbookInitialized(trackSpread)

// spread.fromJSON(jsonData);
var excelIo = new GC.Spread.Excel.IO();
var sheet = spread.getActiveSheet();
document.getElementById('loadExcel').onclick = function() {
    var excelFile = document.getElementById("fileDemo").files[0];
    var password = document.getElementById('password').value;
    // here is excel IO API
    excelIo.open(excelFile, function(json) {
        var workbookObj = json;
        spread.fromJSON(workbookObj);
    }, function(e) {
        // process error
        alert(e.errorMessage);
        if (e.errorCode === 2 /*noPassword*/ || e.errorCode === 3 /*invalidPassword*/ ) {
            document.getElementById('password').onselect = null;
        }
    }, {
        password: password
    });
};
document.getElementById('saveExcel').onclick = function() {

    var fileName = document.getElementById('exportFileName').value;
    var password = document.getElementById('password').value;
    if (fileName.substr(-5, 5) !== '.xlsx') {
        fileName += '.xlsx';
    }

    var json = spread.toJSON();

    // here is excel IO API
    excelIo.save(json, function(blob) {
        saveAs(blob, fileName);
    }, function(e) {
        // process error
        console.log(e);
    }, {
        password: password
    });

};

document.getElementById('trackPrecedentsCell').onclick = function() {
    let sheet = spread.getActiveSheet();
    let trackType = "Precedents";
    let trackCellInfo = sheet.name() + "*" + sheet.getActiveRowIndex() + "*" + sheet.getActiveColumnIndex() + "*" + Math.random();
    trackCellInfoChanged(trackCellInfo, spread, trackSpread, trackType)
};
document.getElementById('trackDependentsCell').onclick = function() {
    let sheet = spread.getActiveSheet();
    let trackType = "Dependents";
    let trackCellInfo = sheet.name() + "*" + sheet.getActiveRowIndex() + "*" + sheet.getActiveColumnIndex() + "*" + Math.random();
    trackCellInfoChanged(trackCellInfo, spread, trackSpread, trackType)
};
document.getElementById('trackAllCell').onclick = function() {
    let sheet = spread.getActiveSheet();
    let trackType = "Both";
    let trackCellInfo = sheet.name() + "*" + sheet.getActiveRowIndex() + "*" + sheet.getActiveColumnIndex() + "*" + Math.random();
    trackCellInfoChanged(trackCellInfo, spread, trackSpread, trackType)
};
```


#### 关于 SpreadJS
[SpreadJS]( https://www.grapecity.com.cn/developer/spreadjs) 是一款基于 HTML5 的纯前端表格控件，兼容 450 多种 Excel 公式，具备“高性能、跨平台、与 Excel 高度兼容”的产品特性。使用 SpreadJS，可直接在 Angular、 React、 Vue 等前端框架中实现高效的模板设计、在线编辑和数据绑定等功能，为最终用户提供高度类似 Excel 的使用体验。
