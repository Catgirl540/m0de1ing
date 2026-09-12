# 微网储能购电优化——C题复现仓库

本仓库包含 C 题四问的原始数据、官方结果模板、模型代码、评价结果和论文图表。四个程序均使用固定参数和确定性算法，可从原始附件重新生成结果。

## 运行环境

- Python 3.12 或更高版本
- 依赖见 `requirements.txt`

```bash
python -m venv .venv
python -m pip install -r requirements.txt
```

Windows PowerShell 激活虚拟环境：

```powershell
.venv\Scripts\Activate.ps1
```

## 顺序复现

在仓库根目录依次运行：

```bash
python one.py
python two.py
python three.py
python four.py
```

运行前请关闭 Excel 中已打开的 `C题/result*.xlsx`，否则 Windows 可能因文件占用而拒绝覆盖。

## 输入与输出

- `C题/附件/`：赛事原始附件与官方空白结果模板
- `one.py`：问题一，生成 `C题/result1.xlsx`
- `two.py`：问题二，生成 `C题/result2.xlsx` 及训练/验证/测试评价
- `three.py`：问题三，生成 `C题/result3.xlsx` 及 MPC 策略评价
- `four.py`：问题四，生成 `C题/result4-2.xlsx`、`C题/result4-3.xlsx` 及电价预测评价
- `C题/论文图表/`：论文使用的核心图表
- `C题/*.md`：各问数学模型与结果说明

只希望重新计算 CSV 指标、不覆盖第四问 Excel 结果时，可运行：

```bash
python four.py --skip-workbooks
```

贝叶斯优化属于耗时实验，正式复现默认直接使用已经冻结的最优参数。如需重新寻优：

```bash
python two.py --bayes-optimize
python four.py --bayes-optimize
```

## 数据集划分

第二至第四问采用“训练、训练、训练、验证、训练、测试”的循环时间划分。参数选择仅使用训练集与验证集，测试集不参与调参。
