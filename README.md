# WPS PPT转图片版

用WPS渲染普通页面为整页图片，保留视频/音频页及原始媒体，并将输出PPTX压缩到指定体积（默认100MB以内）。

## 作为Codex技能安装

将本仓库复制到 `~/.codex/skills/wps-ppt-to-images/`，刷新技能列表后调用 `$wps-ppt-to-images`。

## 运行依赖

Windows、WPS演示（支持 `KWPP.Application` COM接口）、Python 3，以及 `Pillow`、`lxml`。

```powershell
python -m pip install Pillow lxml
powershell -NoProfile -File scripts/render_wps.ps1 -Source 'C:\path\input.pptx' -ImageDirectory 'C:\path\build\images'
python scripts/package_images.py 'C:\path\input.pptx' 'C:\path\build\images' 'C:\path\output.pptx' --max-mb 100
```

使用新输出路径和空图片目录。普通页的动画和可编辑对象会被静态图片取代；视频/音频所在页整体保留。视频不转码。旧式PPT需先用WPS另存临时PPTX。详细流程与检查要求见 [SKILL.md](SKILL.md)。

汇总技能仓库：[Skillls](https://github.com/Melon1234123/Skillls)。仓库不包含演示文稿、视频或用户凭据。
