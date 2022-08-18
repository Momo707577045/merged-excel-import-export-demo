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
