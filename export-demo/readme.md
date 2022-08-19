### 项目功能
- 任意层级合并单元格复杂表头导出

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/001.png)

- 表头与数据项直接映射，无需维护 Excel 索引项匹配关系

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/002.png)

- 自动计算、生成表头合并单元格配置信息

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/003.png)



### [在线示例](https://blog.luckly-mjw.cn/tool-show/merged-excel-import-export-demo/export-demo/index.html)
- 步骤零：如需快速测试，可点击顶部的示例按钮，可快速填充各层级合并单元格 Excel 测试数据

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/004.png)

- 步骤一：输入测试数据源，即从后端获取的数据数组。
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/005.png)
- 步骤二：输入「Excel 表头结构字符串」与「目标数据结构 key」之间的映射关系数组
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/006.png)
  
  - "key" 为 Excel 表头，每一列的所处层级关系。如「基础信息.年龄」对应的就是 Excel 表在第二列中的关系，第一级是「基础信息」，第二级是「年龄」
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/002.png)
  
  - "value" 为数据源数据结构的层级关系。「baseInfo.age」的意思是，将数据源中 baseInfo.age 这个属性的值，设置到 Excel 表的第二列中。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/007.png)
  
  - 【特别注意一】数组的顺序，即为 Excel 表头的顺序。第 N - 1 索引项的配置，即为 Excel 中第 N 列的菜单配置
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/002.png)
  
  - 【特别注意二】Excel 表的层级映射关系与目标对象的层级映射关系，没有强制约束和要求。如Excel 中的 "手机号"，只有一层结构，但完全可以转化为目标对象中的 "contact.phone" 二级结构。反之亦然。
  
  ![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/008.png)
  
- 步骤三：点击「导出 Excel」按钮，即可完成数据 Excel 文件导出。


### 解析后的页面介绍
![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/003.png)

- 左侧区域为数据源转化为 Excel 数据后，每个单元格的数据。包括顶部表头数据，及具体数据行。
  
- 右侧区域为 Excel 表头数据的合并单元格配置项。利用该配置数据，可以通过「xlsx-style」库完成对表头的单元格合并。
  

### 具体函数使用方式
- 示例遵循最少知识原则，项目中的 core.js 文件，即为转换函数所在文件。里面一共不到 200 行代码。可以直接粘贴复制到所需项目里面。

![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/009.png)

- 示例直接使用的 script 文件整体引入的方式。所以迁移到基于 npm 的项目中时，需要将 core.js 中的 ‘Window.XLS.write’ 全局变量使用方式，改为通过 `import { write } from 'xlsx'` 的使用方式。

- 如果本项目确实有帮助到小伙伴，小伙伴有需求的话，可以在 github 中提 issues，有需要的话，将补充基于 npm 的版本，以及带上 ts 类型约束的版本。




### [全部核心源码](https://github.com/Momo707577045/merged-excel-import-export-demo/tree/master/export-demo)

```
/**
 * 将数据源，转化为 Excel 单元格数据，并生成 Excel 表头
 * @param dataList 数据源
 * @param textKeyMaps // Excel 中文表头层级与数据源英文层级间的映射表
 * @param headerFirstRow // 表头首行所在行，为了兼容表格顶部还插入有其他 Excel 行的情况，即表格不在首行
 * @returns {
    headerMerges, // 表头合并单元格配置项
    cells, // 表头及数据项的 Excel 单元格数组
  }
 */
function transformDataToSheetCells(dataList, textKeyMaps, headerFirstRow = 0) {

  // 获取从 textKeyMaps 解析，拆分后的，中英文 keys 数组
  function getKeysList(textKeyMaps) {
    const chineseKeysList = []
    const englishKeysList = []
    textKeyMaps.forEach(textKeyMap => {
      const keyStr = Object.values(textKeyMap)[0]
      const textStr = Object.keys(textKeyMap)[0]
      englishKeysList.push(keyStr.split('.'))
      chineseKeysList.push(textStr.split('.'))
    })
    return {
      englishKeysList,
      chineseKeysList
    }
  }

  // 获取表头行数
  function getHeaderRowNum(chineseKeysList) {
    let maxLevel = 1
    chineseKeysList.forEach(chineseKeys => {
      maxLevel = Math.max(chineseKeys.length, maxLevel)
    })
    return maxLevel
  }

  // 获取表头行 cell 数据
  function getHeaderRows(headerRowNum, chineseKeysList) {
    const headerRows = []
    // 初始化，全部设置为 ''
    for (let rowIndex = 0; rowIndex < headerRowNum; rowIndex++) {
      const row = new Array(chineseKeysList.length).fill('')
      headerRows.push(row)
    }
    // 将表头 cell 设置为对应的中文
    chineseKeysList.forEach((chineseKeys, colIndex) => {
      for (let rowIndex = 0; rowIndex < chineseKeys.length; rowIndex++) {
        headerRows[rowIndex][colIndex] = chineseKeys[rowIndex]
      }
    })

    // 去除需要合并单元格的每一列中。重复的 cell 数据，重复的，则设置为 ''
    headerRows.forEach(headerRow => {
      let lastColValue = ''
      headerRow.forEach((cell, colIndex) => {
        if (lastColValue !== cell) {
          lastColValue = cell
        } else {
          headerRow[colIndex] = ''
        }
      })
    })

    return headerRows
  }

  // 获取合并单元格配置
  function getMerges(headerRowNum, chineseKeysList) {
    const merges = []
    // 竖向合并
    chineseKeysList.forEach((chineseKeys, colIndex) => {
      // 当前列，每一行都有数据，这无需要竖向合并
      if (chineseKeys.length === headerRowNum) {
        return
      }
      // 否则。存在数据需要竖向合并，竖向合并的行数，即为比最高行数少的行数
      merges.push({
        s: {
          r: chineseKeys.length - 1 + headerFirstRow,
          c: colIndex,
        },
        e: {
          r: headerRowNum - 1 + headerFirstRow,
          c: colIndex,
        }
      })
    })
    // 横向合并
    for (let rowIndex = 0; rowIndex < headerRowNum; rowIndex++) {
      const rowCells = chineseKeysList.map(chineseKeys => chineseKeys[rowIndex])
      let preCell = '' // 前一个单元格
      let merge = null // 当前合并配置项
      rowCells.forEach((cell, colIndex) => {
        if (preCell === cell) { // 如果二者相同，则证明需要横向合并单元格
          if (!merge) { // merge 不存在，则创建，
            merge = {
              s: {
                r: rowIndex + headerFirstRow,
                c: colIndex - 1
              },
              e: {
                r: rowIndex + headerFirstRow,
                c: colIndex
              }
            }
            merges.push(merge) // 添加一个合并对象
          } else {
            merge.e.c = colIndex // 修改其合并结束列
          }
        } else {
          preCell = cell
          merge = null
        }
      })
    }
    return merges
  }

  // 获取转化数据结构为 Excel 数据行
  function getDataRows(dataList) {
    const dataRows = []
    dataList.forEach(dataItem => {
      const cells = []
      englishKeysList.forEach(keyLevel => {
        const value = keyLevel.reduce((dataItem, key) => dataItem[key] || '', dataItem).toString()
        cells.push(value)
      })
      dataRows.push(cells)
    })
    return dataRows
  }

  const { englishKeysList, chineseKeysList } = getKeysList(textKeyMaps)
  const headerRowNum = getHeaderRowNum(chineseKeysList)
  const headerMerges = getMerges(headerRowNum, chineseKeysList)
  const headerRows = getHeaderRows(headerRowNum, chineseKeysList)
  const dataRows = getDataRows(dataList)

  return {
    headerMerges,
    cells: [...headerRows, ...dataRows],
  }
}

/**
 * 导出为携带样式的 xlsx 文件
 * @param {*} param
 * @param  param.filename 导出的文件名
 * @param  param.worksheet 导出的 sheet 数据
 */
function toStyleXlsx({ filename, worksheet }) {
  const workbook = {
    SheetNames: [filename],
    Sheets: {
      [filename]: worksheet,
    },
  }

  // writeFile(workbook, filename, { bookType: 'xlsx' })
  let wopts = {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  }
  let wbout = window.XLS.write(workbook, wopts) // 使用xlsx-style 写入
  function s2ab(s) {
    let buf = new ArrayBuffer(s.length)
    let view = new Uint8Array(buf)
    for (let i = 0; i !== s.length; ++i) {
      // eslint-disable-next-line no-bitwise
      view[i] = s.charCodeAt(i) & 0xFF
    }
    return buf
  }
  saveAs(new Blob([s2ab(wbout)], { type: '' }), filename)
}
```

  
### [这里还有 Excel 合并单元格复杂表头导入解析示例](https://blog.luckly-mjw.cn/tool-show/merged-excel-import-export-demo/import-demo/index.html)


![](https://upyun.luckly-mjw.cn/Assets/merged-excel-export-demo/010.webp)

