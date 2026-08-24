# 抽取原语 seed 表

> 唯一事实来源是 `scripts/primitives.py` 的 `PRIMITIVES` 注册表（本文档为人读摘要，若不一致以代码为准）。
> 生成目录：`python scripts/primitives.py`（无参时打印 JSON 目录，见 `catalog()`）。
> 动态补充：用户库（`~/.office-doc-grading/primitives.json`）中的 active 原语会与 seed 合并参与映射（§4.5）。
> 计数核对：以下"word 59 / xlsx 22 / pptx 14"为当前 seed 实数，**以 `catalog()` 运行时输出为准**——
> 新增/改名原语后请重跑生成，勿手改这里的数字。

## word（59 个）

| 原语 | 类型 | 单位 | 说明 |
|---|---|---|---|
| font_name | set | - | 作用域内字体名集合（含显式 ascii/hAnsi/eastAsia 与主题字体如 minorEastAsia） |
| font_size | set | 磅 | 作用域内字号集合 |
| font_color | set | 十六进制 | 作用域内字体颜色集合（如 {FF0000}） |
| highlight | bool | - | 作用域内是否有突出显示 |
| bold / italic / underline | bool | - | 作用域内是否有 |
| align_left/center/right/justify | bool | - | 作用域内是否有该对齐 |
| indent | number | 厘米 | 首行缩进最大值（2字符按字号换算：28pt≈0.99cm；只判"至少一段达标"） |
| indent_all | bool | - | 作用域内所有非空段落是否都有首行缩进（整体性要求，如"正文首行缩进2字符"） |
| line_spacing | number | 磅/倍数 | 行距最大值（固定值取磅） |
| space_before / space_after | number | 磅 | 段前/段后距最大值（0.5行=6磅，afterLines=50 需 XML 直读） |
| margin_top/bottom/left/right | number | 厘米 | 页边距 |
| paper_size | pair | 厘米 | 纸张 [宽, 高] |
| paper_orientation | string | - | portrait/landscape |
| gutter | number | 厘米 | 装订线 |
| gutter_pos | string | - | 装订线位置：left/right（无 gutterPos 属性时缺省=left） |
| header_text / footer_text | string | - | 首页节页眉/页脚文本 |
| header_highlight | string | - | 页眉突出显示颜色（yellow 等） |
| page_number | bool | - | 有页码（pgNumType / 页脚 PAGE 域） |
| page_number_fmt | string | - | 页码格式 fmt（numberInDash≈"-1-,-2-,-3-"类型） |
| page_number_start | number | 页 | 页码起始值 |
| page_break | bool | - | 有分页符 |
| para_count / char_count | number | 个/字符 | 段落数 / 字符数 |
| table_count | number | 个 | 表格数 |
| table_dim | pair | 行x列 | 第一个表格 [行, 列] |
| table_align | string | - | 第一个表格对齐 |
| image_count / image_size | number/pair | 张/厘米 | 图片数 / 第一张 [宽, 高] |
| image_wrap | set | - | 图片环绕集合：inline/square/tight/top_and_bottom/none/through |
| image_position | set | - | 图片位置集合（H:relativeFrom/值 V:…） |
| image_shadow | bool | - | 图片是否有阴影（outerShdw） |
| comment / comment_count | bool/number | 条 | 批注 |
| formula / formula_count | bool/number | 个 | 公式（OMML） |
| footnote / endnote | bool | - | 脚注/尾注 |
| track_changes | bool | - | 修订 |
| watermark | bool | - | 水印 |
| toc | bool | - | 目录 |
| bullet | bool | - | 项目符号/编号（按样式名近似） |
| para_shading | bool | - | 作用域内是否有有效段落底纹（shd 且 fill/themeFill/themeColor 非空） |
| para_border | bool | - | 作用域内是否有有效段落边框（pBdr 存在且非全 none） |
| widow_control | bool | - | 作用域内是否设置孤行控制（段落 pPr 含 w:widowControl 非 val=0） |
| top_line_punct | bool | - | 作用域内是否允许行首标点压缩（段落 pPr 含 w:topLinePunct） |
| page_border | bool | - | 任一节是否设有效页面边框（pgBorders 存在且至少一边 val 非 none） |
| column_two | bool | - | 任一节是否多栏（cols_num>=2）；目标段落归属节精确判定必要时走 LLM |
| column_sep | bool | - | 任一多栏节是否含分隔线（cols_num>=2 且 sep=1） |

## xlsx（22 个）

| 原语 | 类型 | 说明 |
|---|---|---|
| sheet_count | number | 工作表数量 |
| xlsx_freeze | bool | 冻结窗格 |
| xlsx_formula / xlsx_formula_count | bool/number | 公式（抽样 50x30） |
| xlsx_protected | bool | 工作表保护 |
| xlsx_merged | bool | 合并单元格 |
| xlsx_merge_count | number | 合并单元格数量 |
| xlsx_table | bool | 表格样式 |
| xlsx_cond_format | bool | 条件格式 |
| xlsx_cond_type | set | 条件格式类型集合（cellIs/dataBar/colorScale/iconSet 等） |
| xlsx_filter | bool | 筛选 |
| xlsx_font_size | set | 抽样字号 |
| xlsx_col_width / xlsx_row_height | number | 抽样列宽/行高最大值 |
| xlsx_align | set | 抽样单元格水平对齐集合（left/center/right 等） |
| xlsx_border | bool | 是否有单元格边框 |
| xlsx_fill | bool | 是否有单元格底纹/填充色 |
| xlsx_wrap_text | bool | 是否有自动换行单元格 |
| xlsx_number_format | set | 抽查单元格数字格式集合（0.00%/#,##0 等） |
| xlsx_data_validation | bool | 是否有数据验证 |
| xlsx_print_area | bool | 是否设置打印区域 |
| xlsx_hyperlink | bool | 是否有单元格超链接 |

## pptx（14 个）

| 原语 | 类型 | 说明 |
|---|---|---|
| slide_count | number | 幻灯片数 |
| pptx_font_size | set | 各页字号集合 |
| pptx_align_center | bool | 居中对齐文本 |
| pptx_text_align | set | 各页文本对齐集合（left/center/right/justify） |
| pptx_bullet | bool | 项目符号 |
| pptx_animation | bool | 动画（timing 元素） |
| pptx_transition | bool | 切换效果 |
| pptx_notes | bool | 演讲者备注 |
| pptx_hyperlink | bool | 超链接 |
| pptx_image_count | number | 图片总数 |
| pptx_slide_has_title | bool | 是否有含文本的标题占位符 |
| pptx_aspect_ratio | string | 宽高比：16:9 / 4:3 |
| pptx_layout | set | 使用的幻灯片版式集合（Title/Title and Content/Blank 等） |
| pptx_shape_count | number | 形状总数 |

## scope 支持（§4.6）

- 段落类原语（font_*/bold/align_*/indent/line_spacing/space_*）：支持 title/body/text/para scope。
- widow_control / top_line_punct：支持 title/body/text/para scope（按段落 pPr 直读）。
- page_border / column_two / column_sep：节级检测，scope 通常忽略（存在性判定）；段落归属节精确判定走 LLM。
- 表格类（table_*）：支持 table/cell scope。
- image_*：支持 image scope。
- comment_*：支持 comment scope。
- formula_*：支持 formula scope。
- header_text/footer_text/page_number：支持 header/footer scope。
- page scope：仅 first/last（分节近似），精确页码走 LLM 回落。
- 原语未声明支持某 scope 时，agent 应视为不可确定性定位，记 uncovered 或交 LLM。
