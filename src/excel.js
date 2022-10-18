const fs = require('fs');
const xlsx = require('node-xlsx');

function getDepartmentIds(text = '', departments = {}) {
  return text
    .split('\n')
    .map((line) =>
      line
        .trim()
        .split('/')
        .map((t) => departments[t] || '未知ID')
        .join('/')
    )
    .join(' \n');
}

function chunkArr(arr, size) {
  // 判断如果不是数组(就没有length)，或者size没有传值，size小于1，就返回空数组
  if (!arr.length || !size || size < 1) return [];
  let [start, end, result] = [null, null, []];
  for (let i = 0; i < Math.ceil(arr.length / size); i += 1) {
    start = i * size;
    end = start + size;
    result.push(arr.slice(start, end));
  }
  return result;
}

const fields = [
  {
    key: 'name',
    label: '姓名',
    label_en: 'Name',
    example: ''
  },
  {
    key: 'phoneCountryCode',
    label: '手机地区号',
    label_en: 'phoneCountryCode',
    example: ''
  },
  {
    key: 'phone',
    label: '手机号',
    label_en: 'phone',
    example: ''
  },
  {
    key: 'email',
    label: '邮箱',
    label_en: 'email',
    example: ''
  },
  {
    key: 'username',
    label: '用户名',
    label_en: 'username',
    example: ''
  },
  {
    key: 'password',
    label: '密码',
    label_en: 'password',
    example: ''
  },
  {
    key: 'org',
    label: '所属部门',
    label_en: 'org',
    example: ''
  },
  // {
  //   key: 'mainDepartment',
  //   label: '负责部门',
  //   label_en: 'mainDepartment',
  //   example: ''
  // },
  {
    key: 'externalId',
    label: '原系统 ID（externalId)',
    label_en: 'externalId',
    example: ''
  },
  {
    key: 'gender',
    label: '性别',
    label_en: 'gender',
    example: ''
  },
  {
    key: 'country',
    label: '国家',
    label_en: 'country',
    example: ''
  },
  {
    key: 'province',
    label: '省份',
    label_en: 'province',
    example: ''
  },
  {
    key: 'city',
    label: '城市',
    label_en: 'city',
    example: ''
  },
  {
    key: 'isLocked',
    label: '账号状态',
    label_en: 'isLocked',
    example: ''
  }
].map((x) => x.label);

exports.exportExcel = (users) => {
  const originData = users.map((user) => [
    user.name,
    user.phone ? '+86' : null,
    user.phone ?? null,
    user.email ?? null,
    user.name,
    null,
    user.departments ?? null,
    // user.department_leader ? user.departments : null,
    user.externalId ?? null,
    user.gender ?? null,
    null,
    null,
    null,
    '正常'
  ]);
  const data = chunkArr(originData, 1000);
  for (let i = 0; i < data.length; i += 1) {
    data[i].unshift(
      [
        `填写须知：
    请勿修改表格结构；直接上传该文件。`
      ],
      ...Array(13).fill([]),
      fields
    );
    const buffer = xlsx.build([
      {
        name: 'mySheetName',
        data: data[i],
        options: {
          '!cols': Array(fields.length).fill({ wpx: 80 }),
          '!merges': [
            {
              s: { c: 0, r: 0 },
              e: { c: 16, r: 13 }
            }
          ]
          // A1: {
          //   s: {
          //     alignment: {
          //       wrapText: true, // 自动换行
          //       vertical: 'top', // 文字置顶
          //       horizontal: 'left' // 文字向左对齐
          //     }
          //   }
          // }
        }
      }
    ]);

    fs.writeFileSync(`output${i + 1}.xlsx`, buffer, { flag: 'w' });
  }
};
