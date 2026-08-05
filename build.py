#!/usr/bin/env python3
# 构建脚本：把 mapping.js / pipeline.js / app.js 内联进 index.template.html，
# 生成部署用单文件 index.html（GitHub Pages 直接服务此文件）。
#
# 为什么要内联：
#   1) 主程序完全自包含，无需额外请求；
#   2) 快照功能（doSnapshot）直接序列化当前页面生成自包含 HTML，
#      因此页面里的逻辑脚本必须是内联的，否则下载/双击打开的快照会因
#      缺少外部 mapping.js/pipeline.js/app.js 而无法运行。
#   vendor/xlsx.full.min.js 保留为外链（主程序在线加载即可；
#   快照为只读视图不会用到 xlsx，生成时会被剥离）。
#
# 用法：python build.py
import io
import os

HERE = os.path.dirname(os.path.abspath(__file__))


def read(name):
    with io.open(os.path.join(HERE, name), encoding="utf-8") as f:
        return f.read()


template = read("index.template.html")
mapping = read("mapping.js")
pipeline = read("pipeline.js")
app = read("app.js")

html = template
html = html.replace(
    '<script src="mapping.js"></script>',
    "<script>\n" + mapping + "\n</script>",
)
html = html.replace(
    '<script src="pipeline.js"></script>',
    "<script>\n" + pipeline + "\n</script>",
)
html = html.replace(
    '<script src="app.js"></script>',
    "<script>\n" + app + "\n</script>",
)

out = os.path.join(HERE, "index.html")
with io.open(out, "w", encoding="utf-8") as f:
    f.write(html)

print("build ok ->", out, "(%d bytes)" % len(html))
