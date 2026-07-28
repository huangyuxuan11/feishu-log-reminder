# 客户经理日志未填提醒（GitHub Actions 云端版）

每天**工作日（周一~周五）北京时间 20:00 / 20:30 / 20:45**，自动检查飞书两张多维表格，
算出当天还没填「客户经理日志」的人，通过**飞书群机器人**把名单推到你手机。

整个任务跑在 **GitHub 的服务器**上，**不依赖你本地电脑开机**，正好解决
WorkBuddy 本地自动化"别关电脑"的限制。

---

## 一、准备飞书自建应用（读表用）

脚本需要无头访问飞书开放 API（本地 `lark-cli` 扫码登录在云端跑不了），所以要建一个自建应用：

1. 打开 [飞书开放平台](https://open.feishu.cn/) → 开发者后台 → 创建**企业自建应用**。
2. 记下 **App ID** 和 **App Secret**（后续填进 GitHub Secrets）。
3. 权限管理里开通（至少需要其一）：
   - `bitable:app`（旧版，覆盖多维表格读写），或
   - 新版细粒度 `base:record:read` + `base:block:read`。
4. **把该应用添加为两张表的协作者**（否则 API 读不到数据）：
   - 客户经理名单：`DiFmbGm87aaPkxsmz6xctzWTnbe` / 表 `tblQ3UVHeHiS8nZL`
   - 客户经理日志统计：`OteIbzKYha9jKKsprhwcKovNnCh` / 表 `tblAQ9ZGEkAmBgTT`
   - 在表格页面「... → 添加协作者 / 权限」里搜这个应用名加进去即可。
5. 发布应用（或开启"可用范围"），让应用处于启用状态。

> 如果同事已有可复用的飞书应用，直接借用它的 App ID/Secret 也行，跳过 1–5。

## 二、准备飞书群机器人（推送用）

1. 在飞书建一个群（可以只拉你自己 + 这个机器人），群设置 → 群机器人 → 添加机器人 → 自定义机器人。
2. 复制 **Webhook 地址**（形如 `https://open.feishu.cn/open-apis/bot/v2/hook/xxxx`）。
3. **开启签名校验**，复制 **密钥 Secret**（强烈建议开启，防止他人冒发）。
4. 把机器人 webhook 和 secret 填进 GitHub Secrets。

## 三、推到 GitHub 并配置 Secrets

```bash
cd feishu-log-reminder
git init
git add .
git commit -m "init: 客户经理日志未填提醒"
git remote add origin <你的仓库地址>
git push -u origin main
```

在 GitHub 仓库 **Settings → Secrets and variables → Actions → New repository secret** 添加 4 个：

| Secret 名 | 值 |
|---|---|
| `FEISHU_APP_ID` | 自建应用 App ID |
| `FEISHU_APP_SECRET` | 自建应用 App Secret |
| `FEISHU_WEBHOOK` | 群机器人 Webhook 地址 |
| `FEISHU_WEBHOOK_SECRET` | 群机器人签名密钥 |

## 四、验证

- 进仓库 **Actions** 标签页，手动 **Run workflow** 跑一次，看日志是否打印出"已填/未填"结果、手机飞书是否收到消息。
- 确认无误后，定时任务会在工作日 20:00/20:30/20:45（北京时间）自动触发。
- 注意：GitHub 定时任务使用 **UTC**，且高负载时可能延迟几分钟，属正常现象。

## 五、常见问题

- **读不到表 / 401**：应用没被加为表协作者，或 `bitable:app` 权限未开/未发布。
- **收不到消息**：Webhook 地址或签名密钥填错；或机器人被移出群。
- **全员已填也发消息**：是的，脚本会在全员已填时发"✅ 全员已填"，方便确认任务正常跑。

## 文件说明

- `check_unfilled.py` — 主脚本：读两张表 → 算差集 → 推飞书。
- `.github/workflows/remind.yml` — 定时触发（工作日 20:00/20:30/20:45 BJT）。
- `requirements.txt` — 仅依赖 `requests`。
