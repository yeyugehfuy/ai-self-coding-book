# 《方糖AI自编程入门》V1.0

![](./src/images/cover.jpg)

这本书尝试分享如何用自然语言和AI写出真正具有商业价值的复杂应用，而不是那些贪吃蛇玩具。[可以看看这个证据](https://www.bilibili.com/video/BV1BswgeWEkK/?vd_source=fade59d07328dbcb9a0988b7ce98b49d)。


## 第四阶段：个人知识库备份工具架构

本仓库现在补充了一个可扩展的个人知识库备份工具骨架，目标不是只做一次性的导出脚本，而是逐步演进为「个人知识资产同步中心」。

核心设计：

- 平台模块化：石墨、Notion、飞书、浏览器网页等输入平台分别放在 `personal_kb_backup/platforms/`。
- 格式模块化：Word、PDF、Markdown、HTML、Obsidian 等输出格式分别放在 `personal_kb_backup/formats/`。
- 登录模块化：账号登录、Cookie、Token、浏览器会话统一放在 `personal_kb_backup/auth/` 与 `personal_kb_backup/browser/`。
- 导出模块化：批量导出、命名规则、跳过未修改文件、备份报告分别由 `exporters/` 与 `reporting/` 承担。
- 主流程稳定：`personal_kb_backup/core/` 只负责编排，不绑定任何单一平台或格式。

终极目标是普通用户只需要使用根目录的 `同步我的知识库.bat` 和 `更新工具.bat`：前者自动读取配置、登录、同步、增量导出 Word/PDF，并生成备份报告；后者自动更新程序并保留 `.shimo-browser-profile`、`config.toml`、`backup.json` 等用户数据。当前已提供 `config.example.toml` 配置样例、同步入口和更新入口，详细架构见 [第四阶段模块架构](./docs/module-architecture.md)。

## 作者信息

- 作者
    - [Easy](https://ftqq.com)
    - Email: <easychen@gmail.com>
    - 微博：<https://weibo.com/easy>
    - X：<https://x.com/easychen>

## 授权说明

本书采用[CC-BY-NC-SA协议](https://creativecommons.org/licenses/by-nc-sa/4.0/deed.zh-hans)发布。

- 您可以复制、发行、展览、表演、放映、广播或通过信息网络传播本作品，但必须署名作者并添加链接到[本书GitHub仓库](https://github.com/easychen/ai-self-coding-book)。
- 不得为商业目的而使用本作品。
- 仅在遵守与本作品相同的许可条款下，您才能散布由本作品产生的派生作品。

## 电子书

- 可使用 mdbook-epub 工具自行编译：`mdbook-epub --standalone true` 然后 epub 在 book 目录下
- 在官方网站下载(页面最下方)：<https://ft07.com/ai-self-coding-quick-start/>

## 在线阅读

（推荐，体验更好，且可以评论）

1. [时代变了，大人](https://ft07.com/ai-self-coding-times-have-changed/)
1. [用以致学：先用后学，边用边学](https://ft07.com/ai-self-coding-learning-by-using/)
1. [反馈螺旋：小步快跑](https://ft07.com/ai-self-coding-eedback-spiral-small-steps-fast-run/)
1. [时间机器：版本管理](https://ft07.com/ai-self-coding-time-machine/)
1. [质量控制：全量测试与自动化测试](https://ft07.com/ai-self-coding-quality-control-and-testing/)
1. [PDTAC循环：最佳实践](https://ft07.com/ai-self-coding-pdtac/)

## 视频课程（免费）

1. [方糖AI自编程入门 ① AI编程的常见误区和PDTAC循环](https://www.bilibili.com/video/BV1f2cqeFEaA/)
1. [方糖AI自编程入门 ② PDTAC循环开发流程实践](https://www.bilibili.com/video/BV1accBeDEbU/)
2. [方糖AI自编程入门 ③ Web+全平台打包实战：Star Search 语义搜索应用](https://www.bilibili.com/video/BV1FMwseaEkE/)
