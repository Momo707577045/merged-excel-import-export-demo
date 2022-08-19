![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/001.png)


### 项目功能
- 任意层级合并单元格复杂表头解析

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo(https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/002.png)

- 自动转化为目标层级的数据结构

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/003.png)

- 自动生成基于 antdv 的 table 列配置数据 columns 及对于数据源 dataSource。在页面端复现 Excel 效果。

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/004.png)


### [在线示例](https://blog.luckly-mjw.cn/tool-show/merged-excel-import-export-demo/import-demo/index.html)
- 步骤零：如需快速测试，可点击顶部的示例按钮，可快速填充测试数据，并自动下载对应的 Excel 文件，点击上传 Excel 文件即可复现整个使用流程

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/005.png)

- 步骤一：输入「Excel 表头结构字符串」与「目标数据结构 key」之间的映射关系
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/006.png)
  
  - "key" 为 Excel 表头，每一列的所处层级关系。如「基础信息.年龄」对应的就是 Excel 表在第二列中的关系，第一级是「基础信息」，第二级是「年龄」
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/007.png)
  
  - "value" 为需要转换的目标数据结构的层级关系。「baseInfo.age」的意思是，将 Excel 表第二行的数据，转化为目标对象中.baseInfo.age 这个属性。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/008.png)
  
  - 【特别注意一】"key":"value" 的映射关系，没有顺序的要求，无需要按 Excel 表的每一列的数据进行排序。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/009.png)
  
  - 【特别注意二】Excel 表的层级映射关系与目标对象的层级映射关系，没有强制约束和要求。如Excel 中的 "手机号"，只有一层结构，但完全可以转化为目标对象中的 "contact.phone" 二级结构。反之亦然。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/010.png)
  
- 步骤二：点击右侧，上传对应的 Excel 文件，即可完成 Excel 解析。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/011.png)
  

### 解析后的页面介绍

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/013.png)

- 顶部为基于 antdv 的 table 组件的复现效果。复现 Excel 表格中，合并完单元格后的效果。 table 组件用到的 columns 配置，及 dataSource 数据，均由解析函数一并返回。无需要开发者二次开发及维护。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/012.png)
  
- 底部为解析函数解析后返回的三个数据结果。分别为
  - 解析后的 antdv columns 表格列的配置项，直接传递给 table 组件的 columns 属性使用。
  - 解析后的 antdv dataSource 表格数据源，直接传递给 table 组件的 data-source 属性使用。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/004.png)
  
  - 解析后的目标数据结构数组。即根据步骤一设置的映射表，将 Excel 各个单元格数据，转换后的目标数据结构。一般情况下，该数据结构，即为传递给后端的数据结构。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/003.png)
  


### 背景，解决什么问题
- 直接使用 SheetJS 的 XLSX.utils.sheet_to_json 函数进行 excel 数据转化时，仅支持一行表头的 Excel 表格数据解析（只识别 Excel 内容的第一行作为标题），无法识别表头进行过合并单元格的 Excel 数据解析

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/014.png)

- sheet_to_json 转化的 JSON 数据，是以中文为 key 的对象，不符合编程习惯，需要开发者手动进行数据中英文 key 转换，自行转换为目标数据结构。

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/015.png)

- excel 表头的组织与复现，进行 Excel 表上传时，通常需要在前端展现表格内容，给用户需要数据复现及确认。这个过程需要开发者手动组织完成。
- 最少知识原则，提供最小化 demo 示例，没有脚手架，无需要安装 npm 依赖包。纯 html + js 文件。快速测试代码可行性。边改边测进行二次开发。

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/019.png)



### 具体函数使用方式
- 示例遵循最少知识原则，项目中的 core.js 文件，即为转换函数所在文件。里面一共不到 200 行代码。可以直接粘贴复制到所需项目里面。

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/016.png)

- 示例直接使用的 script 文件整体引入的方式。所以迁移到基于 npm 的项目中时，需要将 core.js 中的 ‘XLSX.utils’ 全局变量使用方式，改为通过 `import { utils } from 'xlsx'` 的使用方式。

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/017.png)

- 如果本项目确实有帮助到小伙伴，小伙伴有需求的话，可以在 github 中提 issues，有需要的话，将补充基于 npm 的版本，以及带上 ts 类型约束的版本。

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/018.png)


### 转换函数运行逻辑
- 获取 Excel 所有单元格数据
  - 借用 SheetJS 的 'encode_cell' 方法，及 'format_cell' 方法，遍历获取每个 sheet 中，每个单元格的数据。

```
/**
 * 获取所有单元格数据
 * @param sheet sheet 对象
 * @returns 该 sheet 所有单元格数据
 */
function getSheetCells(sheet) {
  if (!sheet || !sheet['!ref']) {
    return []
  }
  const range = XLSX.utils.decode_range(sheet['!ref'])

  let allCells = []
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; ++rowIndex) {
    let newRow = []
    allCells.push(newRow)
    for (let colIndex = range.s.c; colIndex <= range.e.c; ++colIndex) {
      const cell = sheet[XLSX.utils.encode_cell({
        c: colIndex,
        r: rowIndex
      })]
      let cellContent = ''
      if (cell && cell.t) {
        cellContent = XLSX.utils.format_cell(cell)
      }
      newRow.push(cellContent)
    }
  }
  return allCells
}
```

- 根据 map 映射表，获取 Excel 表头行数
  - 遍历 map 中的每个中文 key，根据'点'关键字拆分层级，找到 Excel 最高层级。从而得到 Excel 表头的行数。区分那些行是表头，那些行是数据。
  
```
  // 获取菜单项在 Excel 中所占行数
  function getHeaderRowNum(textKeyMap) {
    let maxLevel = 1 // 最高层级
    Object.keys(textKeyMap).forEach(textStr => {
      maxLevel = Math.max(maxLevel, textStr.split('.').length)
    })
    return maxLevel
  }
  const headerRowNum = getHeaderRowNum(textKeyMap)
```

- 通过容器，递归解析 Excel 表头配置树装数据结构
  - 通过 lastHeaderLevelColumns 遍历，保存在 Excel 中的每一行的最新容器。
  - 当表头遍历到非空单元格时，往上一级的容器中，加入本层级配置。
  - 从而递归解析，得到 table 表头层级结构数组。
- 使用索引映射表，将 Excel 表头对象格与 Excel 所在行索引进行挂钩
  - 在递归解析表头层级结构时，通过 columnIndexObjMap，记录当前表头数据，所在列。 

```
let headerColumns = [] // 收集 table 组件中，表头 columns 的对象数组结构
  const lastHeaderLevelColumns = [] // 最近一个 columns，用于收集单元格子表头的内容
  const textValueMaps = [] // 以中文字符串为 key 的对象数组，用于收集表格中的数据
  const columnIndexObjMap = [] // 表的列索引，对应在对象中的位置，用于后续获取表格数据时，快速定位每一个单元格

  for (let colIndex = 0; colIndex < headerRows[0].length; colIndex++) {
    const headerCellList = headerRows.map(item => item[colIndex])
    headerCellList.forEach((headerCell, headerCellIndex) => {
      // 如果当前单元格没数据，这证明是合并后的单元格，跳过其处理
      if (!headerCell) {
        return
      }
      const tempColumn = { title: headerCell }

      columnIndexObjMap[colIndex] = tempColumn // 通过 columnIndexObjMap 记录每一列数据，对应到那个对象中，实现一个映射表

      // 如果表头数据第一轮就有值，这证明这是新起的一个表头项目，往 headerColumns 中，新加入一条数据
      if (headerCellIndex === 0) {
        headerColumns.push(tempColumn)
        lastHeaderLevelColumns[headerCellIndex] = tempColumn // 记录当前层级，最新的一个表格容器，可能在下一列数据时合并单元格，下一个层级需要往该容器中添加数据
      } else { // 不是第一列数据，这证明是子项目，需要加入到上一层表头的 children 项，作为其子项目
        lastHeaderLevelColumns[headerCellIndex - 1].children = lastHeaderLevelColumns[headerCellIndex - 1].children || []
        lastHeaderLevelColumns[headerCellIndex - 1].children.push(tempColumn)
        lastHeaderLevelColumns[headerCellIndex] = tempColumn // 记录当前层级的容器索引，可能再下一层级中使用到
      }
    })
  }
```
  
- 利用该索引表，遍历 Excel 数据表的每一行，快速生成每行 Excel 的数据结构
  - 利用 columnIndexObjMap，遍历 Excel 数据表的每一行，往 headerColumns 配置中，插入 value 值，将其设置为特定行，对应列的数据。
  - 通过 Object.create 从 headerColumns 中生成一个对象副本。

```
// 将以数组形式记录的对象信息，转化为正常的对象结构
  function transformListToObj(listObjs) {
    const resultObj = {}
    listObjs.forEach(item => {
      if (item.value) {
        resultObj[item.title] = item.value
        return
      }

      if (item.children && item.children.length > 0) {
        resultObj[item.title] = transformListToObj(item.children)
      }
    })
    return resultObj
  }

  // 以 headerColumns 为对象结构模板，遍历 excel 表数据中的所有数据，并利用 columnIndexObjMap 的映射，快速填充每一项数据
  dataRows.forEach(dataRow => {
    dataRow.forEach((value, index) => {
      columnIndexObjMap[index].value = value
    })
    const titleObj = Object.create(headerColumns) // columnIndexObjMap 的指针索引，仅指向 headerColumns，以 headerColumns 为数据模板，复制一份数据，获得数据填充后的效果对象
    textValueMaps.push(transformListToObj(titleObj)) // 将 listObj 对象转化化普通的对象
  })
```
- 将 JSON 数据结构，进行 key map 映射解析，生成目标数据结构

```
/**
 * 将以点拼接的扁平字符串对象，解析为具有深度的对象
 * @param dotStrObj 点拼接的扁平化字符串对象
 * @returns 具有深度的对象
 */
function parseDotStrObjToObj(dotStrObj) {
  const resultObj = {}
  Object.keys(dotStrObj).forEach(key => {
    let keys = key.split('.')
    keys.reduce((resultObj, curValue, currentIndex) => {
      resultObj[curValue] = currentIndex === keys.length - 1 ? dotStrObj[key] : resultObj[curValue] || {}
      return resultObj[curValue]
    }, resultObj)
  })
  return resultObj
}

/**
 * 将具有深度的对象扁平化，变成以点拼接的扁平字符串对象
 * @param targetObj 具有深度的对象
 * @returns 扁平化后，由点操作符拼接的对象
 */
function transformObjToDotStrObj(targetObj) {
  const resultObj = {}
  function transform(currentObj, preKeys) {
    Object.keys(currentObj).forEach(key => {
      if (Object.prototype.toString.call(currentObj[key]) !== '[object Object]') {
        resultObj[[...preKeys, key].join('.')] = currentObj[key]
      } else {
        transform(currentObj[key], [...preKeys, key])
      }
    })
  }
  transform(targetObj, [])
  return resultObj
}
// 将以中文为 key 的对象，通过 textKeyMap 映射，找到对应的 key，转化为以 key 对键的对象，转化为后端对应的 json 对象
  function transformTextToKey(textDataList, textKeyMap) {
    const textDotStrDataList = textDataList.map(obj => transformObjToDotStrObj(obj))
    let textDotStrDataListStr = JSON.stringify(textDotStrDataList)
    Object.keys(textKeyMap).forEach(text => {
      const key = textKeyMap[text]
      textDotStrDataListStr = textDotStrDataListStr.replaceAll(`"${text}"`, `"${key}"`) // 在这里，通过字符串的替换，实现表头数据层级结构，与实际对象将数据结构的转换
    })
    const keyDotStrDataList = JSON.parse(textDotStrDataListStr)
    const keyDataList = keyDotStrDataList.map(keyDotStrData => parseDotStrObjToObj(keyDotStrData))
    return keyDataList
  }
```

- 返回给 antdv 复现 table 用的 columns 配置，dataSource 表格数据，以及 dataList 后端 JSON 数据

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-import-export-demo/020.png)

- *上述源码确实不太好懂，不太好描述，如果本项目确实能帮助到小伙伴，而小伙伴对源码也感兴趣的话。可以提 issues，我再后补运行逻辑详解及配图*

### 全部核心源码

```
/**
 * 将以点拼接的扁平字符串对象，解析为具有深度的对象
 * @param dotStrObj 点拼接的扁平化字符串对象
 * @returns 具有深度的对象
 */
function parseDotStrObjToObj(dotStrObj) {
  const resultObj = {}
  Object.keys(dotStrObj).forEach(key => {
    let keys = key.split('.')
    keys.reduce((resultObj, curValue, currentIndex) => {
      resultObj[curValue] = currentIndex === keys.length - 1 ? dotStrObj[key] : resultObj[curValue] || {}
      return resultObj[curValue]
    }, resultObj)
  })
  return resultObj
}

/**
 * 将具有深度的对象扁平化，变成以点拼接的扁平字符串对象
 * @param targetObj 具有深度的对象
 * @returns 扁平化后，由点操作符拼接的对象
 */
function transformObjToDotStrObj(targetObj) {
  const resultObj = {}
  function transform(currentObj, preKeys) {
    Object.keys(currentObj).forEach(key => {
      if (Object.prototype.toString.call(currentObj[key]) !== '[object Object]') {
        resultObj[[...preKeys, key].join('.')] = currentObj[key]
      } else {
        transform(currentObj[key], [...preKeys, key])
      }
    })
  }
  transform(targetObj, [])
  return resultObj
}

/**
 * 获取所有单元格数据
 * @param sheet sheet 对象
 * @returns 该 sheet 所有单元格数据
 */
function getSheetCells(sheet) {
  if (!sheet || !sheet['!ref']) {
    return []
  }
  const range = XLSX.utils.decode_range(sheet['!ref'])

  let allCells = []
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; ++rowIndex) {
    let newRow = []
    allCells.push(newRow)
    for (let colIndex = range.s.c; colIndex <= range.e.c; ++colIndex) {
      const cell = sheet[XLSX.utils.encode_cell({
        c: colIndex,
        r: rowIndex
      })]
      let cellContent = ''
      if (cell && cell.t) {
        cellContent = XLSX.utils.format_cell(cell)
      }
      newRow.push(cellContent)
    }
  }
  return allCells
}

/**
 * 获取表头任意层级单元格合并后的表格内容解析
 * @param sheet 一个 sheet 中所有单元格内容
 * @param textKeyMap 表头中文与对应英文 key 之间的映射表
 * @returns antdv 中的表格 column，dataSource，以及转化后的，直接传输给后端的 json 对象数组
 */
function getSheetHeaderAndData(sheet, textKeyMap) {
  // 获取菜单项在 Excel 中所占行数
  function getHeaderRowNum(textKeyMap) {
    let maxLevel = 1 // 最高层级
    Object.keys(textKeyMap).forEach(textStr => {
      maxLevel = Math.max(maxLevel, textStr.split('.').length)
    })
    return maxLevel
  }
  const headerRowNum = getHeaderRowNum(textKeyMap)

  const headerRows = sheet.slice(0, headerRowNum)
  const dataRows = sheet.slice(headerRowNum)

  let headerColumns = [] // 收集 table 组件中，表头 columns 的对象数组结构
  const lastHeaderLevelColumns = [] // 最近一个 columns，用于收集单元格子表头的内容
  const textValueMaps = [] // 以中文字符串为 key 的对象数组，用于收集表格中的数据
  const columnIndexObjMap = [] // 表的列索引，对应在对象中的位置，用于后续获取表格数据时，快速定位每一个单元格

  for (let colIndex = 0; colIndex < headerRows[0].length; colIndex++) {
    const headerCellList = headerRows.map(item => item[colIndex])
    // eslint-disable-next-line no-loop-func
    headerCellList.forEach((headerCell, headerCellIndex) => {
      // 如果当前单元格没数据，这证明是合并后的单元格，跳过其处理
      if (!headerCell) {
        return
      }
      const tempColumn = { title: headerCell }

      columnIndexObjMap[colIndex] = tempColumn // 通过 columnIndexObjMap 记录每一列数据，对应到那个对象中，实现一个映射表

      // 如果表头数据第一轮就有值，这证明这是新起的一个表头项目，往 headerColumns 中，新加入一条数据
      if (headerCellIndex === 0) {
        headerColumns.push(tempColumn)
        lastHeaderLevelColumns[headerCellIndex] = tempColumn // 记录当前层级，最新的一个表格容器，可能在下一列数据时合并单元格，下一个层级需要往该容器中添加数据
      } else { // 不是第一列数据，这证明是子项目，需要加入到上一层表头的 children 项，作为其子项目
        lastHeaderLevelColumns[headerCellIndex - 1].children = lastHeaderLevelColumns[headerCellIndex - 1].children || []
        lastHeaderLevelColumns[headerCellIndex - 1].children.push(tempColumn)
        lastHeaderLevelColumns[headerCellIndex] = tempColumn // 记录当前层级的容器索引，可能再下一层级中使用到
      }
    })
  }

  // 运行以上代码，得到 headerColumns，以及 headerColumns 中，每个对象对应在表格中的哪一行索引

  // 将以数组形式记录的对象信息，转化为正常的对象结构
  function transformListToObj(listObjs) {
    const resultObj = {}
    listObjs.forEach(item => {
      if (item.value) {
        resultObj[item.title] = item.value
        return
      }

      if (item.children && item.children.length > 0) {
        resultObj[item.title] = transformListToObj(item.children)
      }
    })
    return resultObj
  }

  // 以 headerColumns 为对象结构模板，遍历 excel 表数据中的所有数据，并利用 columnIndexObjMap 的映射，快速填充每一项数据
  dataRows.forEach(dataRow => {
    dataRow.forEach((value, index) => {
      columnIndexObjMap[index].value = value
    })
    const titleObj = Object.create(headerColumns) // columnIndexObjMap 的指针索引，仅指向 headerColumns，以 headerColumns 为数据模板，复制一份数据，获得数据填充后的效果对象
    textValueMaps.push(transformListToObj(titleObj)) // 将 listObj 对象转化化普通的对象
  })


  // 根据表头的 title 值，从 textKeyMap 中寻找映射关系，设置 headerColumn 对应的 dataIndex
  function setHeaderColumnDataIndex(headerColumns, preTitle) {
    headerColumns.forEach(headerObj => {
      if (headerObj.children) {
        headerObj.children = setHeaderColumnDataIndex(headerObj.children, [...preTitle, headerObj.title])
      } else {
        const titleStr = [...preTitle, headerObj.title].join('.')
        headerObj.dataIndex = textKeyMap[titleStr]
      }
    })
    return headerColumns
  }

  // 将以中文为 key 的对象，通过 textKeyMap 映射，找到对应的 key，转化为以 key 对键的对象，转化为后端对应的 json 对象
  function transformTextToKey(textDataList, textKeyMap) {
    const textDotStrDataList = textDataList.map(obj => transformObjToDotStrObj(obj))
    let textDotStrDataListStr = JSON.stringify(textDotStrDataList)
    Object.keys(textKeyMap).forEach(text => {
      const key = textKeyMap[text]
      textDotStrDataListStr = textDotStrDataListStr.replaceAll(`"${text}"`, `"${key}"`) // 在这里，通过字符串的替换，实现表头数据层级结构，与实际对象将数据结构的转换
    })
    const keyDotStrDataList = JSON.parse(textDotStrDataListStr)
    const keyDataList = keyDotStrDataList.map(keyDotStrData => parseDotStrObjToObj(keyDotStrData))
    return keyDataList
  }

  headerColumns = setHeaderColumnDataIndex(headerColumns, [])
  const dataList = transformTextToKey(textValueMaps, textKeyMap)
  const dataSource = dataList.map(row => transformObjToDotStrObj(row)) // 将 JSON 对象转化为 “点.” 拼接的扁平对象，使得数据与 headerColumn 中的 dataIndex 相对应。实现 table 的数据填充

  return {
    headerColumns,
    dataList,
    dataSourceList: dataSource,
  }
}
```





