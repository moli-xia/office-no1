# -*- coding: utf-8 -*-
"""
材料排版工具（市委办公室材料排版基本格式）
用法：
  图形界面：双击运行，或  py app.py
  命令行：  py app.py --cli 输入.docx [选项]
"""
import os
import sys
import queue
import argparse
import subprocess

# 字体优先级：Windows 常用公文字体在前，国产 Linux（deepin/统信/麒麟自带方正、
# 思源等字体）的常见名称在后；未命中时默认仍写入规范名称，由 Word/WPS 替换显示
PREFERRED_TITLE = ['方正小标宋简体', '方正小标宋_GBK', '华文中宋', '宋体',
                   'FZXiaoBiaoSong-B05S', 'Noto Serif CJK SC']
PREFERRED_FS = ['仿宋_GB2312', '仿宋', 'FZFangSong-Z02S', 'FangSong_GB2312',
                'Noto Serif CJK SC']
PREFERRED_KAI = ['楷体_GB2312', '楷体', 'FZKai-Z03S', 'KaiTi_GB2312',
                 'AR PL UKai CN', 'Noto Serif CJK SC']
FONT_KEYWORDS = ('宋', '黑', '楷', '仿', '书', '圆', '隶', 'Song', 'Hei', 'Kai',
                 'Fang', 'YaHei', 'Ming', 'Serif', 'CJK', 'FZ', 'Uming', 'UKai',
                 'WenQuanYi')


def _open_folder(path):
    """跨平台打开文件夹（资源管理器 / 文件管理器）"""
    if sys.platform.startswith('win'):
        os.startfile(path)  # noqa: S606
    else:
        subprocess.Popen(['xdg-open', path],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _pick_default(preferred, installed):
    for f in preferred:
        if f in installed:
            return f
    return preferred[0]


ALLOWED_EXTS = ('.docx', '.doc', '.wps')


# ══════════════════════════ 图形界面 ══════════════════════════
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, font as tkfont
    import gongwen

    root = None
    dnd = None
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        root = TkinterDnD.Tk()
        dnd = (TkinterDnD, DND_FILES)
    except Exception:  # noqa: BLE001 —— 未安装 tkinterdnd2 时降级为普通窗口
        root = tk.Tk()
    root.title('材料排版工具 · GB/T 9704—2012 / 市委办公室材料排版基本格式')
    root.geometry('680x700')
    root.minsize(640, 640)
    try:
        ttk.Style().theme_use('vista')
    except tk.TclError:
        pass

    installed = set(tkfont.families())

    def combo_values(preferred):
        names = [f for f in preferred if f in installed]
        names += sorted(f for f in installed
                        if f not in names and any(k in f for k in FONT_KEYWORDS))
        names += [f for f in preferred if f not in names]
        return names

    # ── 文件选择 ──
    filefrm = ttk.LabelFrame(root, text=' Word/WPS 文档（.docx / .doc / .wps）—— 可直接拖入窗口 ')
    filefrm.pack(fill='x', padx=12, pady=(10, 6))
    pathvar = tk.StringVar()
    ttk.Entry(filefrm, textvariable=pathvar).pack(
        side='left', fill='x', expand=True, padx=8, pady=8)

    def browse():
        f = filedialog.askopenfilename(
            title='选择要排版的文档',
            filetypes=[('Word/WPS 文档', '*.docx *.doc *.wps'),
                       ('所有文件', '*.*')])
        if f:
            pathvar.set(f)
    ttk.Button(filefrm, text='浏览…', command=browse).pack(side='left', padx=(0, 8), pady=8)

    # ── 格式标准（标签页）：GB/T 9704—2012 默认，市委办公室格式可选 ──
    nb = ttk.Notebook(root)
    nb.pack(fill='x', padx=12, pady=(4, 6))
    gbfrm = ttk.Frame(nb, padding=(10, 8))
    swbfrm = ttk.Frame(nb, padding=(10, 8))
    nb.add(gbfrm, text=' GB/T 9704—2012 公文格式（默认） ')
    nb.add(swbfrm, text=' 市委办公室材料格式 ')
    ttk.Label(gbfrm, justify='left', text=(
        'GB/T 9704—2012 党政机关公文格式：页边距 上3.7/下3.5/左2.8/右2.6cm（天头37mm、订口28mm）；\n'
        '每面22行、每行28字；层次序数 一、黑体／（一）楷体加粗／1.（1）仿宋加粗；\n'
        '成文日期三号仿宋、右空四字，发文机关署名以成文日期为准居中。'
    )).pack(anchor='w')
    ttk.Label(swbfrm, justify='left', text=(
        '市委办公室材料排版基本格式：页边距 上3.9/下3.3/左2.7/右2.7cm；\n'
        '每面22行、每行28字；层次序数 一、黑体／（一）楷体加粗／1.（1）仿宋加粗；\n'
        '日期三号楷体、右空二字，落款署名与日期一同右对齐。'
    )).pack(anchor='w')

    def current_preset():
        return 'swb' if nb.select() == str(swbfrm) else 'gb9704'

    # ── 选项 ──
    optfrm = ttk.LabelFrame(root, text=' 排版选项')
    optfrm.pack(fill='x', padx=12, pady=6)

    ttk.Label(optfrm, text='标题字体（二号）').grid(row=0, column=0, sticky='w', padx=8, pady=4)
    titlevar = tk.StringVar(value=_pick_default(PREFERRED_TITLE, installed))
    ttk.Combobox(optfrm, textvariable=titlevar, width=22,
                 values=combo_values(PREFERRED_TITLE)).grid(row=0, column=1, sticky='w')

    ttk.Label(optfrm, text='正文字体（三号仿宋）').grid(row=1, column=0, sticky='w', padx=8, pady=4)
    fsvar = tk.StringVar(value=_pick_default(PREFERRED_FS, installed))
    ttk.Combobox(optfrm, textvariable=fsvar, width=22,
                 values=combo_values(PREFERRED_FS)).grid(row=1, column=1, sticky='w')

    ttk.Label(optfrm, text='楷体字体（日期/二层标题）').grid(row=2, column=0, sticky='w', padx=8, pady=4)
    kaivar = tk.StringVar(value=_pick_default(PREFERRED_KAI, installed))
    ttk.Combobox(optfrm, textvariable=kaivar, width=22,
                 values=combo_values(PREFERRED_KAI)).grid(row=2, column=1, sticky='w')

    pagenumvar = tk.BooleanVar(value=True)
    ttk.Checkbutton(optfrm, variable=pagenumvar,
                    text='自动编排页码「— 1 —」（四号宋体，单页居右空一字、双页居左空一字）'
                    ).grid(row=3, column=0, columnspan=2, sticky='w', padx=8, pady=2)
    blankvar = tk.BooleanVar(value=True)
    ttk.Checkbutton(optfrm, variable=blankvar,
                    text='删除正文多余空行（段落之间一般不空行）'
                    ).grid(row=4, column=0, columnspan=2, sticky='w', padx=8, pady=2)
    forwardvar = tk.BooleanVar(value=False)
    ttk.Checkbutton(optfrm, variable=forwardvar,
                    text='转发文件类通知（正文改用三号楷体）'
                    ).grid(row=5, column=0, columnspan=2, sticky='w', padx=8, pady=(2, 8))

    # ── 操作 ──
    btnfrm = ttk.Frame(root)
    btnfrm.pack(fill='x', padx=12, pady=(6, 2))
    gobtn = ttk.Button(btnfrm, text='开 始 排 版')
    gobtn.pack(side='left', pady=4)
    openbtn = ttk.Button(btnfrm, text='打开输出文件夹', state='disabled')
    openbtn.pack(side='left', padx=8, pady=4)
    ttk.Button(btnfrm, text='格式依据说明',
               command=lambda: messagebox.showinfo(
                   '格式依据（GB/T 9704—2012 与 市委办公室材料排版基本格式）',
                   SPEC_HELP)).pack(side='right', pady=4)

    # ── 日志 ──
    logfrm = ttk.LabelFrame(root, text=' 处理日志')
    logfrm.pack(fill='both', expand=True, padx=12, pady=(6, 12))
    logtxt = tk.Text(logfrm, height=10, state='disabled', wrap='word',
                     font=('Microsoft YaHei UI', 9), bg='#fafafa')
    logtxt.pack(fill='both', expand=True, padx=6, pady=6)

    q = queue.Queue()

    def log(msg):
        q.put(msg)

    def worker(src):
        try:
            q.put(f'▶ 正在排版：{os.path.basename(src)}')
            dst, stats = gongwen.format_file(
                src, None,
                preset=current_preset(),
                title_font=titlevar.get().strip() or None,
                fs_font=fsvar.get().strip() or None,
                kai_font=kaivar.get().strip() or None,
                forwarding=forwardvar.get(),
                page_numbers=pagenumvar.get(),
                remove_blank=blankvar.get(),
                log=log)
            q.put(f'✔ 完成！输出文件：{dst}')
            q.put(('__DONE__', dst))
        except gongwen.ConverterUnavailable as e:
            q.put(f'✘ {e}')
            q.put(('__CONV__', str(e), getattr(e, 'src', None)))
        except Exception as e:  # noqa: BLE001
            q.put(f'✘ 失败：{e}')
            q.put(('__ERR__', str(e)))

    def go():
        src = pathvar.get().strip().strip('"')
        if not src:
            messagebox.showwarning('提示', '请先选择或拖入要排版的文档（.docx / .doc / .wps）')
            return
        if not os.path.isfile(src):
            messagebox.showerror('错误', f'文件不存在：{src}')
            return
        if os.path.splitext(src)[1].lower() not in ALLOWED_EXTS:
            messagebox.showerror('错误', '仅支持 .docx / .doc / .wps 文档')
            return
        gobtn.configure(state='disabled')
        openbtn.configure(state='disabled')
        import threading
        threading.Thread(target=worker, args=(src,), daemon=True).start()

    gobtn.configure(command=go)

    # ── 拖拽支持：拖入 .docx/.doc/.wps 即自动开始排版 ──
    if dnd is not None:
        _, DND_FILES = dnd

        def on_drop(event):
            files = [f for f in root.tk.splitlist(event.data) if os.path.isfile(f)]
            if not files:
                return
            f = files[0]
            if len(files) > 1:
                log(f'· 一次只处理一个文件，已取：{os.path.basename(f)}')
            if os.path.splitext(f)[1].lower() not in ALLOWED_EXTS:
                messagebox.showwarning(
                    '不支持的文件类型',
                    f'{os.path.basename(f)}\n\n本工具支持 .docx / .doc / .wps 文档。')
                return
            pathvar.set(f)
            go()

        root.drop_target_register(DND_FILES)
        root.dnd_bind('<<Drop>>', on_drop)

    def check_done():
        try:
            while True:
                msg = q.get_nowait()
                if isinstance(msg, tuple):
                    kind, val = msg
                    if kind == '__DONE__':
                        openbtn.configure(
                            state='normal',
                            command=lambda v=val: _open_folder(os.path.dirname(v)))
                        gobtn.configure(state='normal')
                        messagebox.showinfo('完成', f'排版完成！\n\n输出文件：\n{val}')
                    elif kind == '__CONV__':
                        gobtn.configure(state='normal')
                        src = val[1]
                        if src and messagebox.askyesno('无法自动转换旧格式', val[0] +
                                                       '\n\n是否现在用 WPS/Word 打开该文件？\n'
                                                       '（打开后请「另存为」.docx，再把 .docx 拖入本工具）'):
                            try:
                                _open_path(src)      # 系统默认程序（WPS）打开
                            except Exception:  # noqa: BLE001
                                pass
                    elif kind == '__ERR__':
                        gobtn.configure(state='normal')
                        messagebox.showerror('失败', val)
                    continue
                logtxt.configure(state='normal')
                logtxt.insert('end', msg + '\n')
                logtxt.see('end')
                logtxt.configure(state='disabled')
        except queue.Empty:
            pass
        root.after(120, check_done)

    check_done()
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        pathvar.set(os.path.abspath(sys.argv[1]))
    root.mainloop()


SPEC_HELP = """本工具支持两种格式标准（界面标签页切换，默认国标）：

═══ GB/T 9704—2012 党政机关公文格式（默认）═══
· 幅面 A4；页边距 上3.7 / 下3.5 / 左2.8 / 右2.6cm
  （天头37mm、订口28mm，版心156×225mm）
· 每面22行、每行28字，撑满版心；文字均为黑色
· 标题：二号小标宋，居中，行距固定值33磅
· 正文：三号仿宋_GB2312（转发类通知可用楷体），固定值28磅，
  段首空2字、回行顶格
· 层次序数：一、（黑体）／（一）（楷体加粗）／
  1.（1）（仿宋加粗）
· 成文日期：三号仿宋、右空四字；署名以成文日期为准居中
· 页码：四号半角宋体数字，左右一字线，距版心下边缘约7mm；
  单页码居右空一字、双页码居左空一字
· 数字/西文：Times New Roman

═══ 市委办公室材料排版基本格式 ═══
· 页边距 上3.9 / 下3.3 / 左2.7 / 右2.7cm，其余版式同上
· 层次序数：一、（黑体）／（一）（楷体加粗）／
  1.（1）（仿宋加粗）
· 日期：三号楷体、右空二字；落款署名与日期一同右对齐

说明：
· 红头、版记、份号等版头版记要素需按公文内容人工编排；
· 多行标题的宝塔形分行需人工断行，工具逐行按标题格式编排；
· 首个非空段及其后短段识别为标题；表格内文字仅统一字体与
  黑色，字号保持原样。"""


# ══════════════════════════ 命令行 ══════════════════════════
def run_cli(a):
    import gongwen
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:  # noqa: BLE001
        pass
    dst, _ = gongwen.format_file(
        a.input, a.out, preset=a.preset,
        title_font=a.title_font, fs_font=a.body_font, kai_font=a.kai_font,
        forwarding=a.forwarding, page_numbers=not a.no_page_num,
        remove_blank=not a.keep_blank, log=print)
    print(f'输出：{dst}')


def _dnd_selfcheck():
    """自检拖拽组件（供打包后验证）：输出 DND_OK / DND_FAIL"""
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES  # noqa: F401
        root = TkinterDnD.Tk()
        root.withdraw()
        root.drop_target_register(DND_FILES)
        root.destroy()
        print('DND_OK')
        return 0
    except Exception as e:  # noqa: BLE001
        print('DND_FAIL:', e)
        return 1


def main():
    ap = argparse.ArgumentParser(
        description='材料排版工具（GB/T 9704—2012 党政机关公文格式 / 市委办公室材料排版基本格式）')
    ap.add_argument('--cli', action='store_true', help='命令行模式')
    ap.add_argument('--dnd-check', action='store_true', help=argparse.SUPPRESS)
    ap.add_argument('input', nargs='?', help='输入文档（.docx / .doc / .wps）')
    ap.add_argument('--out', help='输出路径（默认：原名_排版.docx）')
    ap.add_argument('--preset', choices=['gb9704', 'swb'], default='gb9704',
                    help='格式标准：gb9704=GB/T 9704—2012（默认），swb=市委办公室材料排版基本格式')
    ap.add_argument('--title-font', help='标题字体（默认：方正小标宋简体）')
    ap.add_argument('--body-font', help='正文字体（默认：仿宋_GB2312）')
    ap.add_argument('--kai-font', help='楷体字体（默认：楷体_GB2312）')
    ap.add_argument('--forwarding', action='store_true', help='转发文件类通知（正文用楷体）')
    ap.add_argument('--no-page-num', action='store_true', help='不编排页码')
    ap.add_argument('--keep-blank', action='store_true', help='保留正文空行')
    a = ap.parse_args()

    if a.dnd_check:
        sys.exit(_dnd_selfcheck())
    if a.cli or (a.input and not a.cli and os.environ.get('NO_GUI')):
        if not a.input:
            ap.error('--cli 模式需要提供输入文件')
        run_cli(a)
    else:
        run_gui()


if __name__ == '__main__':
    main()
