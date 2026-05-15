#!/usr/bin/env python3
"""
Word 修订标记转换 — 功能区加载项安装脚本
============================================
在 Word 功能区添加"修订工具"选项卡，按钮点击即可逐条审阅修订标记。
只需在终端运行一次：python3 setup_addin.py

工作原理：
  修改 Normal.dotm，注入自定义功能区 XML，按钮关联 VBA 宏。
"""

import os
import sys
import shutil
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

# ── 自定义功能区 XML ──────────────────────────────────────────────
CUSTOM_UI_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="RevisionToolsTab" label="修订工具">
        <group id="MarkGroup" label="标记转换">
          <button id="btnReview"
                  label="逐条审阅"
                  onAction="ReviewAndMarkRevisions_Simple"
                  screentip="逐条审阅修订标记"
                  supertip="逐条显示每条修订标记，可选择：标记、跳过、全部标记或停止"
                  imageMso="AcceptAndMoveToNext"
                  size="large" />
          <button id="btnClear"
                  label="清除格式"
                  onAction="ClearRevisionFormatting"
                  screentip="清除由本工具添加的标记格式"
                  supertip="清除蓝色下划线和红色删除线格式（修订标记已被接受，此操作不可撤销）"
                  imageMso="RejectAndMoveToNext"
                  size="large" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""

# ── 需要更新内容类型的命名空间 ──────────────────────────────────
CT_CUSTOMUI = 'application/xml'
CT_CUSTOMUI_PART = '/customUI/customUI.xml'

CONTENT_TYPES_EXTRA = f"""  <Default Extension="xml" ContentType="{CT_CUSTOMUI}"/>
  <Override PartName="{CT_CUSTOMUI_PART}" ContentType="{CT_CUSTOMUI}"/>"""

RELS_EXTRA = f"""  <Relationship Id="rIdCustomUI"
                Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
                Target="customUI/customUI.xml"/>"""


def find_normal_dotm():
    """查找 Normal.dotm 的路径"""
    candidates = [
        Path.home() / 'Library/Group Containers/UBF8T346G9.Office/User Content/Templates/Normal.dotm',
        Path.home() / 'Library/Application Support/Microsoft/Office/User Templates/Normal.dotm',
        Path.home() / 'Documents/Microsoft 用户数据/Office 2011 用户模板/Normal.dotm',
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def add_custom_ui(normal_path):
    """向 Normal.dotm 注入自定义功能区 XML"""
    print(f"  Normal.dotm 位置: {normal_path}")

    # 备份
    backup = normal_path.with_suffix('.dotm.bak')
    shutil.copy2(normal_path, backup)
    print(f"  已备份到: {backup}")

    # 在临时目录解压
    with tempfile.TemporaryDirectory() as tmpdir:
        # 解压
        with zipfile.ZipFile(normal_path, 'r') as zf:
            zf.extractall(tmpdir)

        # 1. 创建 customUI/customUI.xml
        custom_ui_dir = os.path.join(tmpdir, 'customUI')
        os.makedirs(custom_ui_dir, exist_ok=True)
        with open(os.path.join(custom_ui_dir, 'customUI.xml'), 'w', encoding='utf-8') as f:
            f.write(CUSTOM_UI_XML)
        print("  已添加 customUI/customUI.xml")

        # 2. 更新 [Content_Types].xml
        ct_path = os.path.join(tmpdir, '[Content_Types].xml')
        ct_updated = False
        if os.path.exists(ct_path):
            tree = ET.parse(ct_path)
            root = tree.getroot()
            ns = '{http://schemas.openxmlformats.org/package/2006/content-types}'

            # 添加 Default 或 Override
            has_override = False
            for child in root:
                if child.tag == f'{ns}Override' and child.get('PartName') == CT_CUSTOMUI_PART:
                    has_override = True
                    break

            if not has_override:
                override = ET.SubElement(root, f'{ns}Override')
                override.set('PartName', CT_CUSTOMUI_PART)
                override.set('ContentType', CT_CUSTOMUI)
                ct_updated = True

            if ct_updated:
                tree.write(ct_path, xml_declaration=True, encoding='utf-8')
                print("  已更新 [Content_Types].xml")

        # 3. 更新 _rels/.rels
        rels_path = os.path.join(tmpdir, '_rels', '.rels')
        if os.path.exists(rels_path):
            tree = ET.parse(rels_path)
            root = tree.getroot()

            has_rel = False
            for child in root:
                if child.get('Target') == 'customUI/customUI.xml':
                    has_rel = True
                    break

            if not has_rel:
                ns_rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
                rel = ET.SubElement(root, 'Relationship')
                rel.set('Id', 'rIdCustomUI')
                rel.set('Type', 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility')
                rel.set('Target', 'customUI/customUI.xml')
                tree.write(rels_path, xml_declaration=True, encoding='utf-8')
                print("  已更新 _rels/.rels")

        # 4. 重新打包
        with zipfile.ZipFile(normal_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(tmpdir):
                for file in files:
                    full_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(full_path, tmpdir)
                    zf.write(full_path, arcname)

    print("  功能区 XML 注入完成")


def main():
    print("=" * 50)
    print("Word 修订标记转换 — 功能区加载项安装")
    print("=" * 50)
    print()

    # 检查 Word 是否关闭
    print("请确保 Word 已经完全关闭（Cmd+Q）")
    input("按 Enter 继续...")

    normal = find_normal_dotm()
    if normal is None:
        print("\n❌ 未找到 Normal.dotm。请手动定位文件后运行：")
        print("   python3 setup_addin.py /path/to/Normal.dotm")
        if len(sys.argv) > 1:
            normal = Path(sys.argv[1])
            if not normal.exists():
                print(f"❌ 文件不存在: {normal}")
                sys.exit(1)
        else:
            sys.exit(1)

    try:
        add_custom_ui(normal)
    except Exception as e:
        print(f"\n❌ 安装失败: {e}")
        print("正在恢复备份...")
        backup = normal.with_suffix('.dotm.bak')
        if backup.exists():
            shutil.copy2(backup, normal)
        sys.exit(1)

    print()
    print("=" * 50)
    print("✅ 功能区按钮添加成功！")
    print()
    print("📋 接下来只需一步 — 导入 VBA 宏：")
    print()
    print("  1. 打开 Word")
    print("  2. 点菜单栏「工具」→「宏」→「宏...」")
    print("     (如果菜单栏没有宏，先点「开发工具」→「宏」)")
    print("  3. 「宏的位置」选择 Normal.dotm")
    print("  4. 在「宏名称」输入任意名字，点「创建」")
    print("  5. 删除自动生成的代码")
    print("  6. 打开 RevisionMarker.bas，全选复制，粘贴进去")
    print("  7. Cmd+S 保存，关闭 VBA 编辑器")
    print("  8. 关闭 Word，重新打开")
    print()
    print("  🎯 重新打开后，功能区会出现「修订工具」选项卡")
    print("     点击「逐条审阅」即可开始使用！")
    print("=" * 50)


if __name__ == '__main__':
    main()
