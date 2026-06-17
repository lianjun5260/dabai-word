# 大白陪你背单词 - 项目记忆

## 项目范围
- 只做手机端用户侧，后台词库管理系统由其他团队负责
- 全量设计，不分V1裁剪
- 覆盖四端：微信小程序 + Android + iOS + 鸿蒙(HarmonyOS NEXT)

## 技术选型（已确定）
- 框架：uni-app x 4.64+ (Vue 3 + TypeScript + UTS)
- 语音评测：腾讯云 SOE（主）+ 讯飞 ISE（备用）
- 状态管理：Pinia
- 支付：微信支付/Apple IAP/Google Play/华为支付

## 后台接口对齐
- 后台SRS版本：v2.50
- 前端查询条件：WHERE status=1 AND total_frequency>0
- 词库 = 按标签(word_tag)筛选出的单词集合
- 标签体系：25个默认标签（基础教育16+大学考试5+留学考试4）
- 单词生命周期：审核中(3)→待上架(0)→已上架(1)→已下架(2)→黑名单(4)

## 学习组件（9种）
LC-01 读单词(录音评价) | LC-02 拼单词 | LC-03 读例句 | LC-04 英译汉(选择)
LC-05 汉译英-直选 | LC-06 汉译英-全拼 | LC-07 汉译英-拼读 | LC-08 听力词 | LC-09 听力句

## 学习模式与调度（v2.1新增）
- 4种学习模式：简单/标准/困难/智能（默认智能）
- 简单/标准/困难：后台配置表写死组件编排
- 智能：V1阶段随机调度（保证维度覆盖），V2阶段数据驱动算法
- 组件业务规则待确认（10项），通过基础交互原型逐项确认

## 付费体系（v2.1新增）
- 会员订阅 + AI积分双重付费
- AI积分用于按次计费的大模型调用（语音评测等）
- 会员享积分消耗折扣

## GitHub Pages 固定部署
- 仓库：https://github.com/lianjun5260/dabai-word
- Pages永久地址：https://lianjun5260.github.io/dabai-word/
- 部署方式：prototype/ 文件同步到 docs/ → git push github main
- 必须在 docs/ 保留 .nojekyll 文件
- 改完代码后流程：修改 prototype/ → 同步到 docs/ → git push github → 用户手机刷新即更新

## 登录策略（v2.1新增）
- 单点登录：同账号仅一设备在线
- 新设备登录踢掉旧设备

## 关键文档
- SRS v2.1: docs/SRS-背单词用户端-v2.0.md（内容已是v2.1）/ .docx(v2.1)
- UI设计规范v1.0: docs/UI设计标准与规范-v1.0.md / .docx
- 交互原型: prototype/index.html（23屏单文件HTML原型，含桌面端页面导航面板）
- 后台SRS: 参考材料/后台功能/背单词系统-单词库管理后台_SRS_v2.50.docx
- 效果图: 参考材料/AI英语词库0114_2/ + 大白陪你背单词_效果图/ + 学习组件0402/
- Git仓库: 项目根目录，分支main，文件变更后需随手 `git add` + `git commit`
- 原型部署: 每次更新原型后部署到CloudStudio并生成二维码（绿色#52C41A）

## 每次改完原型后的默认动作（五步）
1. `git add` + `git commit` — 备份
2. 同步更新 SRS 文档（docs/SRS-背单词用户端-v2.0.md）— 保持需求文档与原型一致
3. 同步 `prototype/index.html` → `docs/index.html`
4. `git push github main` — 发布到 GitHub Pages（永久地址）
5. `preview_url` 本地文件 — 更新右侧预览

## UI设计规范要点
- 主色：#52C41A(品牌绿)，辅色#FF4D4F(错误红)、#FAAD14(警告橙)
- 字体：iOS=SF Pro/PingFang SC，Android=Roboto/Noto Sans SC
- 基准屏幕：375pt(iPhone SE)，间距基于4pt网格
- 底部导航3个Tab：首页/统计/我的
- 学习组件统一布局：导航栏+进度 → 徽章+内容 → 提示文字 → 工具栏(原声/收藏/发音) → 录音按钮 → 跳过提示
- 大白IP贯穿：登录页、首页分享横幅、庆祝弹窗
- **返回按钮统一规范**：左上角返回统一使用 `<`（尖括号），不用 `←`（箭头）、不用「返回」文案。CSS class 用 `.back`，功能调用 `goBack()` 或 `goHome()`
