# 客户经理日志未填提醒（GitHub Actions 云端版）

每天**工作日（周一~周五）北京时间 20:00 / 20:30 / 20:45**，自动检查飞书两张多维表格，
算出当天还没填「客户经理日志」的人，通过**飞书自建应用给你发私信**，推到你手机。

**无需建群、无需群机器人。** 一个自建应用既读表、又发提醒（应用私信）。

整个任务跑在 **GitHub 的服务器**上，**不依赖你本地电脑开机**，正好解决
WorkBuddy 本地自动化"别关电脑"的限制。

---

## 一、准备飞书自建应用（读表 + 发私信用同一个应用）

脚本需要无头访问飞书开放 API（本地 `lark-cli` 扫码登录在云端跑不了），所以要建一个自建应用：

1. 打开 [飞书开放平台](https://open.feishu.cn/) → 开发者后台 → 创建**企业自建应用**。
2. 记下 **App ID** 和 **App Secret**（后续填进 GitHub Secrets）。
3. 权限管理里开通：
   - 读表：`bitable:app`（旧版，覆盖多维表格读写），或新版 `base:record:read` + `base:block:read`；
   - 发私信：`im:message` 和 `im:message:send_as_bot`（发送与接收单聊消息）。
4. **把该应用添加为两张表的协作者**（否则 API 读不到数据）：
   - 客户经理名单：`DiFmbGm87aaPkxsmz6xctzWTnbe` / 表 `tblQ3UVHeHiS8nZL`
   - 客户经理日志统计：`OteIbzKYha9jKKsprhwcKovNnCh` / 表 `tblAQ9ZGEkAmBgTT`
   - 在表格页面「... → 添加协作者 / 权限」里搜这个应用名加进去即可。
5. **把你自己加为应用的可用范围/可见范围**（企业自建应用默认对本租户成员可用；若应用设置了白名单，请确认你的账号在可见范围内），否则应用无法给你发私信。
6. 发布应用（或开启"可用范围"），让应用处于启用状态。

> 如果同事已有可复用的飞书应用（已开 `bitable:app` + `im:message` 权限），直接借用它的 App ID/Secret 也行，跳过 1–6。

## 二、获取你的飞书 open_id（接收提醒的账号）

脚本需要知道把私信发给谁。推荐用 **open_id**（最稳）：

- **方式 A（让助手帮你查）**：把你在飞书的姓名告诉助手，助手用已连接的飞书权限（`lark-cli contact`）解析出你的 open_id。
- **方式 B（自己查）**：在飞书开放平台「API 调试台」或任意能拿到自己 open_id 的地方获取；或直接填你的**工号**用作 `FEISHU_USER_ID`（此时 `FEISHU_USER_OPEN_ID` 留空即可，脚本会自动用 `user_id` 模式）。

## 三、配置 GitHub Secrets

仓库已推上 GitHub（公开）。在仓库 **Settings → Secrets and variables → Actions → New repository secret** 添加：

| Secret 名 | 值 | 说明 |
|---|---|---|
| `FEISHU_APP_ID` | 自建应用 App ID | 必填 |
| `FEISHU_APP_SECRET` | 自建应用 App Secret | 必填 |
| `FEISHU_USER_OPEN_ID` | 你的飞书 open_id | 二选一必填（与 `FEISHU_USER_ID` 至少填一个） |
| `FEISHU_USER_ID` | 你的飞书工号 | 可选；不填 open_id 时用作回退 |

## 四、验证

- 进仓库 **Actions** 标签页，手动 **Run workflow** 跑一次，看日志是否打印出"已填/未填"结果、手机飞书是否收到私信。
- 确认无误后，定时任务会在工作日 20:00/20:30/20:45（北京时间）自动触发。
- 注意：GitHub 定时任务使用 **UTC**，且高负载时可能延迟几分钟，属正常现象。

## 五、常见问题

- **读不到表 / 401**：应用没被加为表协作者，或 `bitable:app` 权限未开/未发布。
- **收不到私信 / 报错无权限发消息**：`im:message` + `im:message:send_as_bot` 权限未开；或你不在应用的可见范围内；或 `FEISHU_USER_OPEN_ID` 填错。
- **全员已填也发消息**：是的，脚本会在全员已填时发"✅ 全员已填"，方便确认任务正常跑。

## 文件说明

- `check_unfilled.py` — 主脚本：读两张表 → 算差集 → 用自建应用发飞书私信。
- `.github/workflows/remind.yml` — 定时触发（工作日 20:00/20:30/20:45 BJT）。
- `requirements.txt` — 仅依赖 `requests`。
