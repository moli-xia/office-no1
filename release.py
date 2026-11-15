# -*- coding: utf-8 -*-
"""创建 GitHub Release 并上传产物（网络恢复后运行：py release.py）。

通过 git credential（凭据管理器 GCM）获取令牌，调用 GitHub REST API：
  1. 创建 release v1.0.0（存在则复用）
  2. 上传资产：dist/材料排版工具.exe、linux-安装包.zip
"""
import json
import os
import subprocess
import sys
import urllib.request
import urllib.parse

REPO = 'moli-xia/office-no1'
TAG = 'v2.0.0'
TITLE = 'v2.0.0 材料排版工具（GB/T 9704—2012 / 市委办公室材料排版基本格式）'
NOTES = '''## 材料排版工具 v2.0.0

将 `.docx` / `.doc` / `.wps` 公文拖入窗口，一键按规范完成全文排版（输出保持原文件格式）。

### 格式标准（界面标签页切换）
- **GB/T 9704—2012 党政机关公文格式（默认）**：页边距 上3.7/下3.5/左2.8/右2.6cm，成文日期三号仿宋右空四字、署名以日期为准居中
- **市委办公室材料排版基本格式**：页边距 上3.9/下3.3/左2.7/右2.7cm，日期三号楷体右空二字

### 版式要点
- A4；**每面恒定 22 行、每行 28 字**（撑满版心，孤行控制自动关闭）
- 标题二号小标宋居中、行距固定值 33 磅、标题下自动空一行
- 正文三号仿宋_GB2312、行距固定值 28 磅、段首空 2 字回行顶格；**段前段后强制归零**
- 层次序数：一、黑体 ／（一）楷体加粗 ／ 1.（1）仿宋加粗；小标题与接排正文同段时自动分体设字
- 落款：同行"单位＋日期"自动拆分为两行，**与正文之间空三行**
- 页码四号半角宋体「— N —」，单页居右空一字、双页居左空一字，距版心下边缘约 7mm
- 数字/西文 Times New Roman；全文黑色

### 下载
| 平台 | 文件 | 说明 |
| --- | --- | --- |
| Windows | `材料排版工具.exe` | 免安装单文件，双击即用；支持拖拽 .docx/.doc/.wps 入窗 |
| 国产 Linux（deepin/统信UOS/麒麟） | `linux-安装包.zip` | 解压后 `./install.sh` 一键安装（含桌面图标），详见包内 README-Linux.md |

命令行：`材料排版工具 --cli 文件.docx [--preset gb9704|swb] [--forwarding]`
'''


def get_token():
    r = subprocess.run(['git', 'credential', 'fill'], input='protocol=https\nhost=github.com\n\n',
                       capture_output=True, text=True, timeout=15)
    for line in r.stdout.splitlines():
        if line.startswith('password='):
            return line.split('=', 1)[1]
    raise SystemExit('未获取到 GitHub 凭据（git credential fill 无输出）')


def api(url, token, data=None, method=None, ctype='application/json', extra_headers=None):
    headers = {'Authorization': f'token {token}',
               'User-Agent': 'office-no1-release'}
    if ctype:
        headers['Content-Type'] = ctype
    if extra_headers:
        headers.update(extra_headers)
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=120) as resp:
        body = resp.read()
        return resp.status, json.loads(body) if body and ctype == 'application/json' else body


def main():
    token = get_token()
    print('· 已获取凭据')

    # 先确保远端有该 tag 的提交（本地已推送）
    # 1) 查找既有 release
    try:
        status, rel = api(f'https://api.github.com/repos/{REPO}/releases/tags/{TAG}', token)
        print('· 复用已有 release')
    except urllib.error.HTTPError:
        status, rel = api(f'https://api.github.com/repos/{REPO}/releases', token,
                          data=json.dumps({'tag_name': TAG, 'name': TITLE,
                                           'body': NOTES, 'draft': False,
                                           'prerelease': False}).encode())
        print('· 已创建 release', rel.get('html_url'))
    rid = rel['id']
    print('· Release:', rel['html_url'])

    # 2) 上传资产（重名先删）
    status, assets = api(f'https://api.github.com/repos/{REPO}/releases/{rid}/assets', token)
    exist = {a['name']: a['id'] for a in assets}
    for path, name in ((r'dist\材料排版工具.exe', '材料排版工具.exe'),
                       ('linux-安装包.zip', 'linux-安装包.zip')):
        if name in exist:
            api(f"https://api.github.com/repos/{REPO}/releases/assets/{exist[name]}",
                token, method='DELETE', ctype=None)
            print(f'· 已删除旧资产 {name}')
        with open(path, 'rb') as f:
            data = f.read()
        status, _ = api(
            f'https://uploads.github.com/repos/{REPO}/releases/{rid}/assets?name={urllib.parse.quote(name)}',
            token, data=data, ctype='application/octet-stream',
            extra_headers={'Content-Type': 'application/octet-stream'})
        print(f'· 已上传 {name}（{len(data)//1024}KB, HTTP {status}）')

    print('✔ 完成:', rel['html_url'])


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
