// 数据源
const data = [{
  "baseInfo": {
    "name": "胡勇",
    "age": 82,
    "gender": "女",
    "height": 164,
    "weight": 60
  },
  "contact": {
    "phone": 15662070182,
    "email": "t.ykxuznd@ioxfkolkl.nt"
  },
  "address": {
    "province": "湖北省",
    "city": "延安市"
  }
},
{
  "baseInfo": {
    "name": "程军",
    "age": 92,
    "gender": "男",
    "height": 180,
    "weight": 46
  },
  "contact": {
    "phone": 15652139730,
    "email": "r.hwwnymbjv@zhfgot.cy"
  },
  "address": {
    "province": "黑龙江省",
    "city": "海南藏族自治州"
  }
},
{
  "baseInfo": {
    "name": "卢明",
    "age": 36,
    "gender": "女",
    "height": 159,
    "weight": 41
  },
  "contact": {
    "phone": 15685470329,
    "email": "b.mwskrg@vsqlc.gm"
  },
  "address": {
    "province": "安徽省",
    "city": "吉安市"
  }
},
{
  "baseInfo": {
    "name": "汤平",
    "age": 45,
    "gender": "女",
    "height": 167,
    "weight": 44
  },
  "contact": {
    "phone": 15694331138,
    "email": "s.glujimxnhr@wpgpyde.bs"
  },
  "address": {
    "province": "宁夏回族自治区",
    "city": "陇南市"
  }
}]

// 只有一行 excel 表头的示例
const example1 = {
  data,
  title: '一级菜单示例',
  textKeyMaps: [
    { 姓名: 'baseInfo.name' },
    {年龄: 'baseInfo.age'},
    {性别: 'baseInfo.gender'},
    {身高: 'baseInfo.height'},
    {体重: 'baseInfo.weight'},
    {手机号: 'contact.phone'},
    {邮箱: 'contact.email'},
    {所在省: 'address.province'},
    {所在市: 'address.city'},
  ]
}


// 两行且存在单元格合并的 excel 表头的示例
const example2 = {
  data,
  title: '二级菜单示例',
  textKeyMaps: [
    {姓名: 'baseInfo.name'},
    {'基础信息.年龄': 'baseInfo.age'},
    {'基础信息.性别': 'baseInfo.gender'},
    {'基础信息.身高': 'baseInfo.height'},
    {'基础信息.体重': 'baseInfo.weight'},
    {手机号: 'contact.phone'},
    {邮箱: 'contact.email'},
    {'地址信息.所在省': 'address.province'},
    {'地址信息.所在市': 'address.city'},
  ],
}

// 三行且存在单元格合并的 excel 表头的示例
const example3 = {
  data,
  title: '三级菜单示例',
  textKeyMaps: [
    {姓名: 'baseInfo.name'},
    {'基础信息.年龄': 'baseInfo.age'},
    {'基础信息.性别': 'baseInfo.gender'},
    {'基础信息.身高': 'baseInfo.height'},
    {'基础信息.体重': 'baseInfo.weight'},
    {'基础信息.联系方式.手机号': 'contact.phone'},
    {'基础信息.联系方式.邮箱': 'contact.email'},
    {'地址信息.所在省': 'address.province'},
    {'地址信息.所在市': 'address.city'},
  ],
}

// 四行且存在单元格合并的 excel 表头的示例
const example4 = {
  data,
  title: '四级菜单示例',
  textKeyMaps: [
    {'成员数据.姓名': 'baseInfo.name'},
    {'成员数据.基础信息.年龄': 'baseInfo.age'},
    {'成员数据.基础信息.性别': 'baseInfo.gender'},
    {'成员数据.基础信息.身高': 'baseInfo.height'},
    {'成员数据.基础信息.体重': 'baseInfo.weight'},
    {'成员数据.基础信息.联系方式.手机号': 'contact.phone'},
    {'成员数据.基础信息.联系方式.邮箱': 'contact.email'},
    {'成员数据.地址信息.所在省': 'address.province'},
    {'成员数据.地址信息.所在市': 'address.city'},
  ],
}