# Office Forge 性能与鲁棒性基线

`benchmarks/run_benchmarks.py` 用真实 Python 引擎子进程执行固定负载，记录墙钟时间与 Windows 峰值工作集，并与 `benchmarks/baseline.json` 中的阈值比较。完整回归（`tests/run_regression.py`，非 `--quick`）会调用该套件；超过阈值即判定 `REGRESSION FAILED`。

## 运行

```bash
# 默认规模：8,000 段 Word / 12,000×10 Excel / 120 页 PPT
"$PYTHON" benchmarks/run_benchmarks.py

# 大文件档：30,000 段 / 50,000×20 / 300 页（手动基准，不进入默认回归）
"$PYTHON" benchmarks/run_benchmarks.py --full
```

输出每个用例一行 `PASS <case> seconds=... peakMB=...`，最后打印 `BENCHMARK PASSED`。`--update-baseline` 会用当前机器实测值重写阈值（仅在确认硬件/运行时变化后手工执行，不要作为日常修复）。阈值默认取实测值的 1.5 倍并保留约 50% 内存余量，避免正常机器波动造成误报。

## 默认基线（本机，2026-11-22）

| 用例 | 实测秒 | 峰值 MB | maxSeconds | maxPeakMB |
|---|---:|---:|---:|---:|
| docx_create（8,000 段 + 200 行表格） | 13.09 | 80.2 | 30.64 | 119.7 |
| docx_read | 16.10 | 82.8 | 35.529 | 123.9 |
| docx_audit（profile） | 15.15 | 82.8 | 36.546 | 124.6 |
| xlsx_create（12,000×10） | 3.82 | 121.9 | 9.25 | 183 |
| xlsx_read | 0.75 | 61.8 | 1.665 | 92.7 |
| xlsx_audit（health） | 2.27 | 156.1 | 4.899 | 234.5 |
| pptx_create（120 页） | 1.29 | 70.6 | 3.117 | 105.8 |
| pptx_read | 0.84 | 75.5 | 1.888 | 113.1 |
| pptx_audit（deck） | 0.83 | 75.3 | 1.894 | 112.9 |
| pdf_corrupt | 0.72 | 60.5 | 1.644 | 90.9 |
| pdf_empty | 0.73 | 60.6 | 1.639 | 90.8 |
| pdf_encrypted_no_password | 0.73 | 60.5 | 1.679 | 90.9 |

## 鲁棒性断言

- `pdf_corrupt`：非 PDF 字节必须返回 `INVALID_PDF`，不得把垃圾字节当成空文档。
- `pdf_empty`：0 字节文件必须返回 `INVALID_PDF`。
- `pdf_encrypted_no_password`：加密 PDF 无密码读取必须返回 `PDF_PASSWORD_REQUIRED`，不得崩溃。

## 说明

- 峰值内存通过 Windows `GetProcessMemoryInfo` 的 `PeakWorkingSetSize` 采样；其他平台该列为 0，仅比较耗时。
- 计时包含 Python 子进程启动、JSON 序列化和引擎初始化，适合发现相对回归，不等同于单函数微基准。
- 大文件档 `--full` 用于交付前人工压测；默认回归保持约一分钟内完成，避免拖慢日常测试。
