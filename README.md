# 科创债 ETF 曲线构建

本项目用于自动更新科创债 ETF 相关数据，构建 ETF 成分券收益率曲线，并与中债中短期票据收益率曲线进行对比，生成图表、利差时间序列和邮件正文素材。

整体流程包括：

1. 刷新本地 Excel 数据源。
2. 合并中债收益率曲线历史数据和更新数据。
3. 从 Wind 获取科创债 ETF 成分券、剩余期限、隐含评级和估值收益率。
4. 分评级构建关键期限收益率点位。
5. 计算 ETF 成分券曲线相对中债曲线的利差。
6. 输出图片和 Excel 附件。
7. 将图片嵌入邮件并发送。

## 项目结构

```text
.
├── main.py                                      # 主入口，串联 step2、step3、step4
├── step1刷新数据.py                              # 刷新本地 Excel / Wind 插件公式数据
├── step2合并数据.py                              # 合并中债曲线历史文件和更新文件
├── step3画图.py                                  # 获取 ETF 成分券、计算曲线和利差、生成图表
├── step4发送邮件.py                              # 收集图片和附件，生成并发送邮件
├── requirements.txt                             # Python 依赖清单，目前为空，需按运行环境补充
├── 运行说明.txt                                  # 简短运行说明
├── 交易日历_更新.xlsx                            # 交易日历和日期控制文件
├── 期限结构_中债AAA_历史.xlsx                     # AAA 中债曲线历史数据
├── 期限结构_中债AAA_更新.xlsx                     # AAA 中债曲线更新数据
├── 期限结构_中债AA+_历史.xlsx                     # AA+ 中债曲线历史数据
├── 期限结构_中债AA+_更新.xlsx                     # AA+ 中债曲线更新数据
├── 科创债ETF成分券vs中债中短票据_利差时间序列.xlsx  # 累积保存的历史利差时间序列
└── 邮件图片文件夹/                                # 每个交易日生成的图片和附件输出目录
```

项目中同时保留了 `.ipynb` 版本，主要用于交互式调试和历史留痕；日常自动化运行以 `.py` 脚本为准。

## 运行环境

建议使用 Windows + Anaconda 环境运行，因为项目依赖 WindPy、Excel、xlwings 和本地 Office 能力。

### 必要软件

- Windows
- Python 3.10 或以上
- Microsoft Excel
- Wind 金融终端及 WindPy
- 可正常使用的邮箱 SMTP 授权码

### Python 依赖

当前 `requirements.txt` 为空，但脚本实际使用以下包：

```text
pandas
numpy
matplotlib
tqdm
scipy
scikit-learn
xlwings
openpyxl
WindPy
```

可按实际环境安装：

```powershell
pip install pandas numpy matplotlib tqdm scipy scikit-learn xlwings openpyxl
```

`WindPy` 通常随 Wind 客户端安装，不能简单通过公共 PyPI 安装。运行前请确认：

```python
from WindPy import w
w.start()
```

能够正常启动并返回可用连接。

## 数据文件说明

### `交易日历_更新.xlsx`

用于控制脚本使用的日期。`step3画图.py` 会读取：

- `上一個交易日`：作为 Wind 数据和图表输出日期。
- `今天`：作为历史回溯计算的结束日期。

如果日期列名在 Excel 中发生变化，需要同步修改脚本中的读取字段。

### 中债曲线文件

项目使用两组中债曲线文件：

- `期限结构_中债AAA_历史.xlsx`
- `期限结构_中债AAA_更新.xlsx`
- `期限结构_中债AA+_历史.xlsx`
- `期限结构_中债AA+_更新.xlsx`

`step2合并数据.py` 会把更新文件中的日期列合并进历史文件。

### 利差时间序列文件

`科创债ETF成分券vs中债中短票据_利差时间序列.xlsx` 用于保存历史利差序列。`step3画图.py` 每次运行后会把新计算的数据追加合并到该文件，并去重保存。

## ETF 范围

`step3画图.py` 当前覆盖以下科创债 ETF：

```text
159200.SZ
159400.SZ
159600.SZ
159700.SZ
511120.SH
551030.SH
551500.SH
551550.SH
551900.SH
```

如需增删 ETF，请修改 `step3画图.py` 中的 `dict_etf` 和历史计算部分的 `ETF_DICT`。

## 曲线构建逻辑

`step3画图.py` 的核心步骤如下：

1. 通过 Wind `w.wset("etfconstituent", ...)` 获取 ETF 成分券。
2. 通过 Wind `w.wss(...)` 获取债券剩余期限、隐含评级、估值收益率等字段。
3. 优先使用行权剩余期限 `termifexercise`，缺失时回退到普通剩余期限 `ptmyear`。
4. 过滤隐含评级，仅保留 `AAA` 和 `AA+`。
5. 过滤永续债，仅保留非永续债。
6. 在关键期限窗口内取样，并使用稳健统计方法计算 ETF 成分券收益率点。
7. 与中债中短期票据曲线在相同关键期限上对齐。
8. 计算利差：

```text
Spread(bp) = (ETF(%) - 中债曲线(%)) * 100
```

当前关键期限为：

```text
1Y, 2Y, 3Y, 4Y, 5Y, 8Y, 9Y, 10Y
```

## 缺失值处理

部分评级和期限可能因为样本不足导致空值，例如 `AA+ 10Y`。脚本中已加入插值处理：

- 当天曲线图绘制前，按期限顺序对关键期限点进行线性插值。
- 历史时间序列保存前，按 `Rating + Tenor` 分组并按日期插值。
- 两端缺值使用最近的有效值补齐。

这样可以减少最终图表和 Excel 中连续空值过多的问题。

## 输出结果

运行后会在以下目录生成结果：

```text
邮件图片文件夹/{上一交易日}/
```

主要输出包括：

- 当日 ETF 成分券估值相关信息 Excel。
- `AAA` 和 `AA+` 成分券曲线对比图。
- 各关键期限的历史利差走势图。
- 可供邮件发送的图片和附件。

同时会更新：

```text
科创债ETF成分券vs中债中短票据_利差时间序列.xlsx
```

## 运行方式

### 推荐方式

在项目根目录运行：

```powershell
python main.py
```

`main.py` 会依次执行：

1. `step2合并数据.py`
2. `step3画图.py`
3. `step4发送邮件.py`

当前 `main.py` 中 `step1刷新数据.py` 的调用被注释掉了。如需先刷新 Excel 数据，可以手动取消注释或单独运行。

### 分步骤运行

刷新 Excel 数据：

```powershell
python step1刷新数据.py
```

合并中债曲线数据：

```powershell
python step2合并数据.py
```

生成图表和利差序列：

```powershell
python step3画图.py
```

发送邮件：

```powershell
python step4发送邮件.py
```

## 邮件配置

邮件发送配置位于 `step4发送邮件.py`：

```python
SMTP_HOST = "smtp.qq.com"
SMTP_PORT = 465
USE_SSL = True
USERNAME = "..."
PASSWORD = "..."
FROM_ADDR = "..."
TO_ADDRS = [...]
CC_ADDRS = [...]
```

上线或共享仓库前，建议把邮箱账号、授权码、收件人列表迁移到环境变量或本地配置文件中，避免敏感信息直接提交到 GitHub。

## 常见问题

### 1. `KeyError: Timestamp(...)`

通常是脚本按某个交易日读取中债曲线列，但历史曲线 Excel 中没有完全匹配的日期列。检查：

- `交易日历_更新.xlsx` 中的上一交易日是否正确。
- `期限结构_中债AAA_历史.xlsx` 和 `期限结构_中债AA+_历史.xlsx` 是否已经合并了对应日期列。
- 日期列是 `Timestamp`、字符串还是 Excel 日期格式。

### 2. Wind 数据获取失败

检查：

- Wind 客户端是否已登录。
- WindPy 是否能正常 `w.start()`。
- 债券代码或 ETF 代码是否有权限访问。
- 查询日期是否为有效交易日。

### 3. 图片中文字显示异常

脚本中设置了：

```python
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
```

如果本机没有 SimHei 字体，图中文字可能显示异常。可安装字体或修改为本机可用中文字体。

### 4. 邮件发送失败

检查：

- SMTP 授权码是否正确。
- 邮箱是否开启 SMTP 服务。
- 端口和 SSL / STARTTLS 设置是否匹配。
- 附件和图片路径是否存在。

