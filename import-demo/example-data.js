// 只有一行 excel 表头的示例
const example1 = {
  title: '一级菜单示例',
  fileUrl: 'https://upyun.luckly-mjw.cn/Assets/excel-import/excel-example-1.xlsx',
  textKeyMap: {
    姓名: 'baseInfo.name',
    年龄: 'baseInfo.age',
    性别: 'baseInfo.gender',
    身高: 'baseInfo.height',
    体重: 'baseInfo.weight',
    手机号: 'contact.phone',
    邮箱: 'contact.email',
    所在省: 'address.province',
    所在市: 'address.city',
  },
}


// 两行且存在单元格合并的 excel 表头的示例
const example2 = {
  title: '二级菜单示例',
  fileUrl: 'https://upyun.luckly-mjw.cn/Assets/excel-import/excel-example-2.xlsx',
  textKeyMap: {
    姓名: 'baseInfo.name',
    '基础信息.年龄': 'baseInfo.age',
    '基础信息.性别': 'baseInfo.gender',
    '基础信息.身高': 'baseInfo.height',
    '基础信息.体重': 'baseInfo.weight',
    手机号: 'contact.phone',
    邮箱: 'contact.email',
    '地址信息.所在省': 'address.province',
    '地址信息.所在市': 'address.city',
  },
}

// 三行且存在单元格合并的 excel 表头的示例
const example3 = {
  title: '三级菜单示例',
  fileUrl: 'https://upyun.luckly-mjw.cn/Assets/excel-import/excel-example-3.xlsx',
  textKeyMap: {
    姓名: 'baseInfo.name',
    '基础信息.年龄': 'baseInfo.age',
    '基础信息.性别': 'baseInfo.gender',
    '基础信息.身高': 'baseInfo.height',
    '基础信息.体重': 'baseInfo.weight',
    '基础信息.联系方式.手机号': 'contact.phone',
    '基础信息.联系方式.邮箱': 'contact.email',
    '地址信息.所在省': 'address.province',
    '地址信息.所在市': 'address.city',
  },
}

// 四行且存在单元格合并的 excel 表头的示例
const example4 = {
  title: '四级菜单示例',
  fileUrl: 'https://upyun.luckly-mjw.cn/Assets/excel-import/excel-example-4.xlsx',
  textKeyMap: {
    '成员数据.姓名': 'baseInfo.name',
    '成员数据.基础信息.年龄': 'baseInfo.age',
    '成员数据.基础信息.性别': 'baseInfo.gender',
    '成员数据.基础信息.身高': 'baseInfo.height',
    '成员数据.基础信息.体重': 'baseInfo.weight',
    '成员数据.基础信息.联系方式.手机号': 'contact.phone',
    '成员数据.基础信息.联系方式.邮箱': 'contact.email',
    '成员数据.地址信息.所在省': 'address.province',
    '成员数据.地址信息.所在市': 'address.city',
  },
}