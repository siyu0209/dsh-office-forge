# Office Forge 性能与鲁棒性基线

`benchmarks/run_benchmarks.py` 用真实 Python 引擎子进程执行固定负载，记录墙钟时间与 Windows 峰值工作集，并与 `benchmarks/baseline.json` 中的阈值比较。完整回归（`tests/run_regression.py`，非 `--quick`）会调用该套件；超过阈值即判定 `REGRESSION FAILED`。

## 运行

```bash
# 默认规模：8,000 段 Word / 12,000×10 Excel / 120 页 PPT
"$PYTHON" benchmarks/run_benchmarks.py

# 大文件档：30,000 段 / 50,000×20 / 300 页（手动基准，不进入默认回归）
"$PYTHON" benchmarks/run_benchmarks.py --full
```

输出每个用例一行 `PASS <case> seconds=... peakMB=...`，最后打印 `BENCHMARK PASSED`。`--update-baseline` 会用当前机器实测值重写阈值（仅在确认硬件/运行时变化后手工执行，不要作为日常修复）。阈值按共享工作站上 2–3 倍实测值设置，并保留约 50% 内存余量；目标是发现数量级回归，不是固定硬件跑分。

## 默认基线（本机，2026-11-24，含环境余量）

| 用例 | 实测秒 | 峰值 MB | maxSeconds | maxPeakMB |
|---|---:|---:|---:|---:|
| docx_create（8,000 段 + 200 行表格） | 35.69 | 80.5 | 70 | 120.8 |
| docx_read | 43.92 | 83.0 | 95 | 124.5 |
| docx_audit（profile） | 31.73 | 83.1 | 100 | 124.6 |
| xlsx_create（12,000×10） | 7.47 | 122.6 | 24 | 183.9 |
| xlsx_read | 1.47 | 62.2 | 5 | 93.3 |
| xlsx_audit（health） | 5.07 | 156.3 | 15 | 234.5 |
| pptx_create（120 页） | 2.56 | 71.0 | 9 | 106.5 |
| pptx_read | 1.89 | 76.0 | 5 | 114.0 |
| pptx_audit（deck） | 1.58 | 75.8 | 4 | 113.7 |
| deck_import（120 页 Deck IR 反向解析） | 1.19 | 77.0 | 4 | 120 |
| pdf_corrupt | 1.45 | 60.9 | 4 | 91.3 |
| pdf_empty | 1.40 | 60.9 | 4 | 91.3 |
| pdf_encrypted_no_password | 1.72 | 60.9 | 4 | 91.3 |

## 鲁棒性断言

- `pdf_corrupt`：非 PDF 字节必须返回 `INVALID_PDF`，不得把垃圾字节当成空文档。
- `pdf_empty`：0 字节文件必须返回 `INVALID_PDF`。
- `pdf_encrypted_no_password`：加密 PDF 无密码读取必须返回 `PDF_PASSWORD_REQUIRED`，不得崩溃。

## 说明

- 峰值内存通过 Windows `GetProcessMemoryInfo` 的 `PeakWorkingSetSize` 采样；其他平台该列为 0，仅比较耗时。
- 计时包含 Python 子进程启动、JSON 序列化和引擎初始化，适合发现相对回归，不等同于单函数微基准。
- 大文件档 `--full` 用于交付前人工压测；默认回归保持约一分钟内完成，避免拖慢日常测试。
