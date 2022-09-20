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

exports.exportExcel = (users, departments) => {
  const data = users.map((user) => [
    null,
    user.externalId ?? null,
    user.email ?? null,
    user.email ? '是' : '否',
    user.phone ?? null,
    null,
    null,
    null,
    'basic:import-excel',
    null,
    0,
    null,
    user.gender ?? 'U',
    null,
    null,
    null,
    null,
    // 注册时间
    null,
    null,
    '否',
    user.departments ?? null,
    user.department_leader ? '是' : '否',
    getDepartmentIds(user.departments, departments)
  ]);
  data.unshift([
    '昵称',
    '用户名',
    '邮箱',
    '邮箱是否验证',
    '手机号',
    'unionid',
    'openid',
    '密码',
    '注册方式',
    '公司',
    '登录次数',
    '最近登录 IP',
    '性别',
    '国家',
    '省份',
    '城市',
    '地址',
    '注册时间',
    '上次登录时间',
    '账号是否被禁用',
    '部门',
    '是否为部门 Leader',
    '部门 ID'
  ]);
  const buffer = xlsx.build([{ name: 'mySheetName', data }]);
  fs.writeFileSync('output.xlsx', buffer, { flag: 'w' });
};
