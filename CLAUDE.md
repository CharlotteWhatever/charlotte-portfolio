# Charlotte Portfolio Website · CLAUDE.md

## Quick Facts
- **项目**：Charlotte Chi 个人品牌/求职网站
- **文件**：`index.html`（单文件 HTML/CSS/JS，零框架）
- **部署**：GitHub Pages (`charlottewhatever.github.io/charlotte-portfolio`) + Cloudflare Pages
- **Git**：✅ 已初始化，有完整版本历史

## 配色 · Dark Explorer（当前方案，待用户确认）

| 色号 | CSS 变量 | 用途 |
|------|----------|------|
| `#1C2A42` | `--night` | 标题、主色调、深色区 |
| `#B8845C` | `--copper` | 亮点、强调、按钮 |
| `#F4EFE6` | `--ivory` | 页面底色 70% |
| `#5A6A76` | `--muted` | 正文 |
| `#FAF7F2` | `--card` | 卡片底色 |
| `#E2D9CC` | `--border` | 边框 |

比例：70% 暖象牙 + 20% 深夜蓝 + 10% 暖铜。

## 页面模块（9 个）

Nav → Hero → Experience（4公司时间线）→ Case Study（CAME）→ Portfolio（4卡片）→ AI Playground（4项目）→ Recommendations（4人）→ Downloads（7项）→ Footer

## 版本历史

- `v1-desktop` (tag `4ad9455`, 2026-06-10)：Neumorphism 蓝，仅桌面端
- `v2-mobile-responsive` (tag `3ff4d90`, 2026-06-11)：4级断点 + AI板块重写

回滚：`git checkout v1-desktop -- index.html`

## 用户偏好（强制规则）

- 方引号「」作为引用标记，不要用 ""
- 线性 icon（Phosphor Icons），不用 emoji
- 左右留适当 padding，不能满版全宽
- 圆角 8-10px
- **不要自己创造设计** → 套用现有设计或找参考
- **文案以原始内容方案为准** → 不要改文案
- 修改已有功能时先确认「替代还是新增」
- 设计探索：最多出 2 个方向，用户选一个再继续

## 关键踩坑

1. **AI 图片拉伸（5轮才解决）**：根因是 Flexbox `min-height: auto`。方案：`position: absolute` + `aspect-ratio` + `object-fit: cover`
2. **汉堡菜单（3轮失败放弃）**：毛玻璃+neumorphism阴影残留无法消除，回退到缩小字号方案
3. **设计探索占40%时间但零产出**：8轮配色探索，最终未采用任何一套。教训：设计对齐在开发前完成，用截图/色板而非完整HTML
4. **Bug修复靠猜测而非诊断**：没打开 DevTools 检查 computed height。正确做法：先量化再修

## 待办

- [ ] 用户确认 Dark Explorer 配色
- [ ] 中英文切换功能实现
- [ ] emoji → Phosphor Icons
- [ ] 照片上传 hero 区域
- [ ] 导出中英文简历 PDF
- [ ] 自定义域名绑定（可选）

## 简历系统规则

- 英文简历 → Gmail（国际送达保障）
- 中文简历 → Foxmail（国内直达）
- JS 语言切换时动态改 `mailto:` 和显示文本
- 邮箱替换前确认：中文版和英文版可能有不同邮箱策略
